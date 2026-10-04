import {
  decryptLegacyRc4,
  decryptRc4CryptoApi,
  parseLegacyRc4Verifier,
  parseRc4CryptoApiHeader,
  verifyLegacyRc4Password,
  verifyRc4CryptoApiPassword,
} from "@office-open/core";
import type { CompoundFileReader } from "@office-open/core";

import { DocParseError } from "../errors";

const FLAG_TABLE_ONE = 0x0200;

export function decryptWordStreams(
  encryptedWord: Uint8Array,
  encryptedTable: Uint8Array,
  reader: CompoundFileReader,
  password?: string,
): { word: Uint8Array; table: Uint8Array; tablePath: string; data?: Uint8Array } {
  if (password === undefined) {
    throw new DocParseError("Encrypted Word documents are not supported");
  }
  const version = new DataView(
    encryptedTable.buffer,
    encryptedTable.byteOffset,
    encryptedTable.byteLength,
  );
  if (encryptedTable.byteLength < 4 || version.getUint16(2, true) !== 0x0002) {
    throw new DocParseError("Encrypted Word documents are not supported");
  }
  const majorVersion = version.getUint16(0, true);
  const decrypt =
    majorVersion === 1
      ? (data: Uint8Array): Uint8Array => {
          const verifier = parseLegacyRc4Verifier(data, 4);
          if (!verifyLegacyRc4Password(password, verifier)) {
            throw new DocParseError("Invalid Word document password");
          }
          return decryptLegacyRc4(data, password, verifier.salt);
        }
      : (data: Uint8Array): Uint8Array => {
          const { keySizeBits, verifier } = parseRc4CryptoApiHeader(data, 4);
          if (!verifyRc4CryptoApiPassword(password, verifier, keySizeBits)) {
            throw new DocParseError("Invalid Word document password");
          }
          return decryptRc4CryptoApi(data, password, verifier.salt, keySizeBits);
        };
  const decryptedWord = decrypt(encryptedWord);
  const decryptedWordCopy = new Uint8Array(decryptedWord);
  decryptedWordCopy.set(encryptedWord.subarray(0, 0x44));
  const flagsOffset = 10;
  decryptedWordCopy[flagsOffset]! &= ~0x80;
  decryptedWordCopy[flagsOffset + 1]! &= ~0x01;
  new DataView(decryptedWordCopy.buffer, decryptedWordCopy.byteOffset).setUint32(14, 0, true);
  const tablePath =
    (readUint16(decryptedWordCopy, 10, "Invalid Word document: FIB is truncated") &
      FLAG_TABLE_ONE) !==
    0
      ? "1Table"
      : "0Table";
  const table = decrypt(encryptedTable);
  const data = reader.entry("Data") ? decrypt(reader.read("Data")) : undefined;
  return { word: decryptedWordCopy, table, tablePath, data };
}

function requireRange(bytes: Uint8Array, offset: number, length: number, message: string): void {
  if (offset < 0 || length < 0 || offset > bytes.byteLength || length > bytes.byteLength - offset) {
    throw new DocParseError(message);
  }
}

function readUint16(bytes: Uint8Array, offset: number, message: string): number {
  requireRange(bytes, offset, 2, message);
  return bytes[offset]! | (bytes[offset + 1]! << 8);
}
