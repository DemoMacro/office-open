/**
 * Password hashing utilities for OOXML document protection.
 *
 * Implements ECMA-376 Agile Encryption password derivation.
 * Passwords are encoded as UTF-16LE per the OOXML specification.
 *
 * In Node.js: uses `node:crypto` for SHA-512 (~11× faster than hash.js).
 * In browsers: falls back to hash.js.
 *
 * @module
 */

import hash from "hash.js";

// ── Native crypto detection (Node.js ESM) ──

type Hasher = { update(data: Uint8Array): Hasher; digest(): Uint8Array };
type CreateHashFn = (algo: string) => Hasher;

let _createHash: CreateHashFn | undefined;

try {
  // top-level await — ESM standard.
  // In Node.js: resolves to crypto module with createHash.
  // In browsers: import("node:crypto") throws — falls back to hash.js.
  const nodeCrypto = await import("node:crypto");
  if (typeof nodeCrypto.createHash !== "function") throw new Error("no native crypto");
  _createHash = nodeCrypto.createHash;
} catch {
  // Browser — hash.js fallback
}

// ── Helpers ──

function toBase64(bytes: Uint8Array): string {
  const binary = String.fromCharCode(...bytes);
  return btoa(binary);
}

function utf16leEncode(str: string): Uint8Array {
  const bytes = new Uint8Array(str.length * 2);
  for (let i = 0; i < str.length; i++) {
    const code = str.charCodeAt(i);
    bytes[i * 2] = code & 0xff;
    bytes[i * 2 + 1] = (code >> 8) & 0xff;
  }
  return bytes;
}

// ── Public API ──

export function randomBytes(length: number): Uint8Array {
  const bytes = new Uint8Array(length);
  // eslint-disable-next-line @typescript-eslint/no-unnecessary-condition
  if (typeof crypto !== "undefined" && crypto.getRandomValues) {
    crypto.getRandomValues(bytes);
  } else {
    for (let i = 0; i < length; i++) {
      bytes[i] = Math.floor(Math.random() * 256);
    }
  }
  return bytes;
}

/**
 * Computes the ECMA-376 Agile Encryption hash for a password.
 *
 * Algorithm: SHA-512(salt + UTF-16LE(password)), then iterate
 * SHA-512(hash + i.toUint32LE()) spinCount times.
 */
export function hashPasswordAgile(password: string, salt: Uint8Array, spinCount: number): string {
  const pwBytes = utf16leEncode(password);

  // h = SHA-512(salt + password_utf16le)
  const initial = new Uint8Array(salt.length + pwBytes.length);
  initial.set(salt, 0);
  initial.set(pwBytes, salt.length);

  let h: Uint8Array;

  if (_createHash) {
    // Node.js fast path (~11× faster)
    h = _createHash("sha512").update(initial).digest();
    for (let i = 0; i < spinCount; i++) {
      const buf = new Uint8Array(h.length + 4);
      buf.set(h, 0);
      buf[h.length] = i & 0xff;
      buf[h.length + 1] = (i >> 8) & 0xff;
      buf[h.length + 2] = (i >> 16) & 0xff;
      buf[h.length + 3] = (i >> 24) & 0xff;
      h = _createHash("sha512").update(buf).digest();
    }
  } else {
    // Browser fallback (hash.js)
    h = hash.sha512().update(initial).digest() as unknown as Uint8Array;
    for (let i = 0; i < spinCount; i++) {
      const counter = new Uint8Array(4);
      new DataView(counter.buffer).setUint32(0, i, true);
      const buf = new Uint8Array(h.length + 4);
      buf.set(h, 0);
      buf.set(counter, h.length);
      h = hash.sha512().update(buf).digest() as unknown as Uint8Array;
    }
  }

  return toBase64(new Uint8Array(h));
}

/**
 * Derives a password hash for OOXML document protection.
 *
 * Returns base64-encoded hashValue, saltValue, spinCount, and algorithmName
 * suitable for direct use in document protection options.
 */
export function derivePasswordHash(
  password: string,
  spinCount = 100000,
): { hashValue: string; saltValue: string; spinCount: number; algorithmName: string } {
  const salt = randomBytes(16);
  const hashValue = hashPasswordAgile(password, salt, spinCount);
  return {
    hashValue,
    saltValue: toBase64(salt),
    spinCount,
    algorithmName: "SHA-512",
  };
}
