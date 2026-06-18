/**
 * Shared base64 encode/decode helpers.
 *
 * Both prefer the native `Uint8Array` base64 methods (Node 22+, modern
 * browsers): they avoid the intermediate binary string that `btoa`/`atob`
 * materialize and cannot stack-overflow on large buffers. Node `Buffer` is the
 * secondary path; a manual loop covers older runtimes.
 *
 * @module
 */

/**
 * Decode a base64 string into a `Uint8Array`.
 *
 * Prefers native `Uint8Array.fromBase64` (no intermediate binary string), then
 * Node `Buffer` (zero-copy), then `atob`.
 */
export function decodeBase64(input: string): Uint8Array {
  const fromBase64 = (Uint8Array as { fromBase64?: (s: string) => Uint8Array }).fromBase64;
  if (typeof fromBase64 === "function") return fromBase64.call(Uint8Array, input);
  if (typeof Buffer !== "undefined") return Buffer.from(input, "base64");
  return Uint8Array.from(atob(input), (c) => c.codePointAt(0)!);
}

/**
 * Encode a `Uint8Array` into a base64 string.
 *
 * Prefers native `Uint8Array.prototype.toBase64` (no intermediate binary
 * string), then Node `Buffer` (zero-copy), then a `btoa` fallback.
 *
 * The fallback builds the binary string in a loop rather than
 * `String.fromCharCode(...bytes)` — the spread form places every byte on the
 * call stack and overflows for large buffers (V8 caps function arguments near
 * ~65k).
 */
export function encodeBase64(bytes: Uint8Array): string {
  const toBase64 = (Uint8Array.prototype as { toBase64?: () => string }).toBase64;
  if (typeof toBase64 === "function") return toBase64.call(bytes);
  if (typeof Buffer !== "undefined") return Buffer.from(bytes).toString("base64");
  let binary = "";
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}
