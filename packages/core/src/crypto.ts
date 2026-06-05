/**
 * Password hashing utilities for OOXML document protection.
 *
 * Implements ECMA-376 Agile Encryption password derivation.
 *
 * @module
 */

import hash from "hash.js";

/**
 * Encodes a Uint8Array to a plain base64 string (no data URI prefix).
 */
function toBase64(bytes: Uint8Array): string {
  // btoa is available in Node.js 16+ and all browsers
  const binary = String.fromCharCode(...bytes);
  return btoa(binary);
}

/**
 * Generates cryptographically secure random bytes.
 */
export function randomBytes(length: number): Uint8Array {
  const bytes = new Uint8Array(length);
  // crypto is available in Node.js 15+ and all modern browsers
  // eslint-disable-next-line @typescript-eslint/no-unnecessary-condition
  if (typeof crypto !== "undefined" && crypto.getRandomValues) {
    crypto.getRandomValues(bytes);
  } else {
    // Fallback for very old environments
    for (let i = 0; i < length; i++) {
      bytes[i] = Math.floor(Math.random() * 256);
    }
  }
  return bytes;
}

/**
 * Computes the ECMA-376 Agile Encryption hash for a password.
 *
 * Algorithm: SHA-512(salt + UTF-8(password)), then iterate SHA-512(hash + i.toUint32LE()) spinCount times.
 */
export function hashPasswordAgile(password: string, salt: Uint8Array, spinCount: number): string {
  // h = SHA-512(salt + password)
  let h = hash.sha512().update(salt).update(new TextEncoder().encode(password)).digest();

  // Iterate: h = SHA-512(h + i as uint32LE)
  for (let i = 0; i < spinCount; i++) {
    const counter = new Uint8Array(4);
    new DataView(counter.buffer).setUint32(0, i, true); // little-endian
    h = hash.sha512().update(h).update(counter).digest();
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
