import { DocParseError } from "../errors";
export function requireRange(
  bytes: Uint8Array,
  offset: number,
  length: number,
  message: string,
): void {
  if (offset < 0 || length < 0 || offset > bytes.byteLength || length > bytes.byteLength - offset) {
    throw new DocParseError(message, {
      part: "stream",
      offset,
      length,
      byteRange: [offset, offset + length],
      reason: "out-of-range",
    });
  }
}

export function readUint16(bytes: Uint8Array, offset: number, message: string): number {
  requireRange(bytes, offset, 2, message);
  return bytes[offset]! | (bytes[offset + 1]! << 8);
}

export function readUint32(bytes: Uint8Array, offset: number, message: string): number {
  requireRange(bytes, offset, 4, message);
  return (
    (bytes[offset]! |
      (bytes[offset + 1]! << 8) |
      (bytes[offset + 2]! << 16) |
      (bytes[offset + 3]! << 24)) >>>
    0
  );
}
