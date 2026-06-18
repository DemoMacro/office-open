/**
 * Binary input normalization.
 *
 * Accepts the full range of binary inputs (Buffer/Uint8Array/ArrayBuffer/
 * DataView/number[]/string/base64 data URL/…) and normalizes to `Uint8Array`.
 * Centralized here so every entry point (packer, patch, descriptor helpers)
 * shares one definition instead of re-declaring per module.
 *
 * @module
 */
import { decodeBase64 } from "./base64";

/** Supported binary input shapes. */
export type DataType =
  | ArrayBufferLike
  | Blob
  | DataView
  | number[]
  | ReadableStream
  | string
  | Uint8Array;

// Matches data:[<mediatype>][;base64],<data> — a base64 data URL. Mirrors the
// `isBase64DataURL` check in unjs/undio so plain strings stay UTF-8 text.
const DATA_URL_RE = /^data:([\w.+-]+\/[\w.+-]+)?;base64,/;

/** Test whether a string is a base64 data URL (`data:[mime];base64,...`). */
export function isBase64DataURL(input: string): boolean {
  return DATA_URL_RE.test(input);
}

/** Normalize any supported binary input to a `Uint8Array`. */
export function toUint8Array(data: DataType): Uint8Array {
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (data instanceof DataView)
    return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  if (typeof data === "string") {
    const match = data.match(DATA_URL_RE);
    if (match) return decodeBase64(data.slice(match[0].length));
    return new TextEncoder().encode(data);
  }
  if (Array.isArray(data)) return new Uint8Array(data);
  if (data instanceof Blob) throw new TypeError("Blob input requires async processing");
  if (data instanceof ReadableStream)
    throw new TypeError("ReadableStream input requires async processing");
  throw new TypeError(`Unsupported data type: ${typeof data}`);
}
