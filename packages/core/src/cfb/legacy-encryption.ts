import { md5, sha1 } from "@noble/hashes/legacy.js";

const BLOCK_KEY_BYTES = 4;
const LEGACY_KEY_SIZE_BITS = 40;
const FULL_KEY_BYTES = 16;
const VERIFIER_BYTES = 16;
const RC4_VERIFIER_HASH_BYTES = 20;

function utf16le(value: string): Uint8Array {
  const result = new Uint8Array(value.length * 2);
  const view = new DataView(result.buffer);
  for (let index = 0; index < value.length; index += 1) {
    view.setUint16(index * 2, value.charCodeAt(index), true);
  }
  return result;
}

function concatBytes(parts: readonly Uint8Array[]): Uint8Array {
  const result = new Uint8Array(parts.reduce((total, part) => total + part.byteLength, 0));
  let offset = 0;
  for (const part of parts) {
    result.set(part, offset);
    offset += part.byteLength;
  }
  return result;
}

function rc4(key: Uint8Array, data: Uint8Array): Uint8Array {
  const state = Array.from({ length: 256 }, (_, index) => index);
  let swap = 0;
  for (let index = 0; index < 256; index += 1) {
    swap = (swap + state[index]! + key[index % key.byteLength]!) & 0xff;
    [state[index], state[swap]] = [state[swap]!, state[index]!];
  }
  const result = new Uint8Array(data.byteLength);
  let left = 0;
  let right = 0;
  for (let index = 0; index < data.byteLength; index += 1) {
    left = (left + 1) & 0xff;
    right = (right + state[left]!) & 0xff;
    [state[left], state[right]] = [state[right]!, state[left]!];
    result[index] = data[index]! ^ state[(state[left]! + state[right]!) & 0xff]!;
  }
  return result;
}

function blockKey(
  derive: (block: number) => Uint8Array,
  block: number,
): { key: Uint8Array; block: number } {
  return { key: derive(block), block: block + 1 };
}

function decryptInBlocks(
  data: Uint8Array,
  blockSize: number,
  derive: (block: number) => Uint8Array,
  initialBlock = 0,
): Uint8Array {
  const result = new Uint8Array(data.byteLength);
  let state = blockKey(derive, initialBlock);
  for (let offset = 0; offset < data.byteLength; offset += blockSize) {
    const decrypted = rc4(
      state.key,
      data.subarray(offset, Math.min(offset + blockSize, data.byteLength)),
    );
    result.set(decrypted, offset);
    state = blockKey(derive, state.block);
  }
  return result;
}

function bytesEqual(left: Uint8Array, right: Uint8Array): boolean {
  return left.byteLength === right.byteLength && left.every((byte, index) => byte === right[index]);
}

/** Derive the 128-byte RC4 key from a legacy password, salt, and block number. */
export function deriveLegacyRc4Key(password: string, salt: Uint8Array, block: number): Uint8Array {
  const passwordHash = md5(utf16le(password));
  const truncatedHash = passwordHash.subarray(0, 5);
  const intermediateInput = new Uint8Array((truncatedHash.byteLength + salt.byteLength) * 16);
  for (let index = 0; index < 16; index += 1) {
    const offset = index * (truncatedHash.byteLength + salt.byteLength);
    intermediateInput.set(truncatedHash, offset);
    intermediateInput.set(salt, offset + truncatedHash.byteLength);
  }
  const intermediate = md5(intermediateInput);
  const blockBytes = new Uint8Array(BLOCK_KEY_BYTES);
  new DataView(blockBytes.buffer).setUint32(0, block, true);
  return md5(concatBytes([intermediate.subarray(0, 5), blockBytes]));
}

/** Derive the RC4 CryptoAPI key from a Unicode password, salt, key size, and block number. */
export function deriveRc4CryptoApiKey(
  password: string,
  salt: Uint8Array,
  keySizeBits: number,
  block: number,
): Uint8Array {
  const initial = sha1(concatBytes([salt, utf16le(password)]));
  const blockBytes = new Uint8Array(BLOCK_KEY_BYTES);
  new DataView(blockBytes.buffer).setUint32(0, block, true);
  const final = sha1(concatBytes([initial, blockBytes]));
  const bytes = keySizeBits === 0 ? LEGACY_KEY_SIZE_BITS : keySizeBits;
  if (bytes === LEGACY_KEY_SIZE_BITS) {
    const key = new Uint8Array(FULL_KEY_BYTES);
    key.set(final.subarray(0, bytes / 8));
    return key;
  }
  return final.subarray(0, bytes / 8);
}

/** Decrypt a legacy RC4 payload, restarting the key at every block boundary. */
export function decryptLegacyRc4(
  data: Uint8Array,
  password: string,
  salt: Uint8Array,
  blockSize = 512,
): Uint8Array {
  return decryptInBlocks(data, blockSize, (block) => deriveLegacyRc4Key(password, salt, block));
}

/** Decrypt an RC4 CryptoAPI payload, restarting the key at every block boundary. */
export function decryptRc4CryptoApi(
  data: Uint8Array,
  password: string,
  salt: Uint8Array,
  keySizeBits: number,
  blockSize = 512,
  initialBlock = 0,
): Uint8Array {
  return decryptInBlocks(
    data,
    blockSize,
    (block) => deriveRc4CryptoApiKey(password, salt, keySizeBits, block),
    initialBlock,
  );
}

interface LegacyVerifier {
  readonly salt: Uint8Array;
  readonly encryptedVerifier: Uint8Array;
  readonly encryptedVerifierHash: Uint8Array;
}

function readVerifier(data: Uint8Array, offset: number, hashBytes: number): LegacyVerifier {
  if (offset + 40 + hashBytes > data.byteLength) {
    throw new Error("Invalid legacy Office encryption verifier");
  }
  return {
    salt: data.slice(offset + 4, offset + 20),
    encryptedVerifier: data.slice(offset + 20, offset + 36),
    encryptedVerifierHash: data.slice(offset + 40, offset + 40 + hashBytes),
  };
}

/** Parse an RC4 CryptoAPI header that immediately follows its four version bytes. */
export function parseRc4CryptoApiHeader(
  data: Uint8Array,
  offset = 0,
): {
  keySizeBits: number;
  verifier: LegacyVerifier;
} {
  if (offset + 8 > data.byteLength) {
    throw new Error("Invalid legacy Office RC4 CryptoAPI header");
  }
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const headerSize = view.getUint32(offset + 4, true);
  const headerEnd = offset + 8 + headerSize;
  if (headerEnd + 4 > data.byteLength) {
    throw new Error("Invalid legacy Office RC4 CryptoAPI header");
  }
  const keySizeBits = view.getUint32(offset + 24, true);
  return { keySizeBits, verifier: readVerifier(data, headerEnd, RC4_VERIFIER_HASH_BYTES) };
}

/** Parse the fixed legacy RC4 verifier that immediately follows its four version bytes. */
export function parseLegacyRc4Verifier(data: Uint8Array, offset = 0): LegacyVerifier {
  if (offset + VERIFIER_BYTES * 3 > data.byteLength) {
    throw new Error("Invalid legacy Office encryption verifier");
  }
  return {
    salt: data.slice(offset, offset + VERIFIER_BYTES),
    encryptedVerifier: data.slice(offset + VERIFIER_BYTES, offset + VERIFIER_BYTES * 2),
    encryptedVerifierHash: data.slice(offset + VERIFIER_BYTES * 2, offset + VERIFIER_BYTES * 3),
  };
}

/** Verify a candidate password against the fixed legacy RC4 verifier. */
export function verifyLegacyRc4Password(
  password: string,
  { salt, encryptedVerifier, encryptedVerifierHash }: LegacyVerifier,
): boolean {
  const decrypted = rc4(
    deriveLegacyRc4Key(password, salt, 0),
    concatBytes([encryptedVerifier, encryptedVerifierHash]),
  );
  return bytesEqual(md5(decrypted.subarray(0, VERIFIER_BYTES)), decrypted.subarray(VERIFIER_BYTES));
}

/** Verify a candidate password against an RC4 CryptoAPI verifier. */
export function verifyRc4CryptoApiPassword(
  password: string,
  { salt, encryptedVerifier, encryptedVerifierHash }: LegacyVerifier,
  keySizeBits: number,
): boolean {
  const key = deriveRc4CryptoApiKey(password, salt, keySizeBits, 0);
  const decrypted = rc4(key, concatBytes([encryptedVerifier, encryptedVerifierHash]));
  return bytesEqual(
    sha1(decrypted.subarray(0, VERIFIER_BYTES)),
    decrypted.subarray(VERIFIER_BYTES),
  );
}
