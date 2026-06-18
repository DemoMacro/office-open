/**
 * Password hashing utilities for OOXML document protection.
 *
 * Implements ECMA-376 Agile Encryption password derivation.
 * Passwords are encoded as UTF-16LE per the OOXML specification.
 *
 * Uses @noble/hashes for SHA-512 — audited, cross-platform,
 * and comparable in speed to node:crypto for this workload.
 *
 * @module
 */

import { sha512 } from "@noble/hashes/sha2.js";

import { encodeBase64 } from "./base64";

// ── Helpers ──

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

  let h = sha512(initial);
  for (let i = 0; i < spinCount; i++) {
    const buf = new Uint8Array(h.length + 4);
    buf.set(h, 0);
    buf[h.length] = i & 0xff;
    buf[h.length + 1] = (i >> 8) & 0xff;
    buf[h.length + 2] = (i >> 16) & 0xff;
    buf[h.length + 3] = (i >> 24) & 0xff;
    h = sha512(buf);
  }

  return encodeBase64(new Uint8Array(h));
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
    saltValue: encodeBase64(salt),
    spinCount,
    algorithmName: "SHA-512",
  };
}
