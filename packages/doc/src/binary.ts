import { DocParseError, type DocParseContext } from "./errors";

export type DocBinaryContext = Partial<Omit<DocParseContext, "format">>;

export interface DocBinaryLocation {
  part: DocParseContext["part"];
  path?: string;
}

export function requireByteRange(
  bytes: Uint8Array,
  offset: number,
  length: number,
  message: string,
  context: DocBinaryContext = {},
  location: DocBinaryLocation = { part: "stream" },
): void {
  if (offset < 0 || length < 0 || offset > bytes.byteLength || length > bytes.byteLength - offset) {
    throw new DocParseError(message, {
      part: location.part,
      ...(location.path ? { path: location.path } : {}),
      offset,
      length,
      byteRange: [offset, offset + length],
      reason: "out-of-range",
      ...context,
    });
  }
}

export function readUint16(
  bytes: Uint8Array,
  offset: number,
  message: string,
  context: DocBinaryContext = {},
  location: DocBinaryLocation = { part: "stream" },
): number {
  requireByteRange(bytes, offset, 2, message, context, location);
  return bytes[offset]! | (bytes[offset + 1]! << 8);
}

export function readUint32(
  bytes: Uint8Array,
  offset: number,
  message: string,
  context: DocBinaryContext = {},
  location: DocBinaryLocation = { part: "stream" },
): number {
  requireByteRange(bytes, offset, 4, message, context, location);
  return (
    (bytes[offset]! |
      (bytes[offset + 1]! << 8) |
      (bytes[offset + 2]! << 16) |
      (bytes[offset + 3]! << 24)) >>>
    0
  );
}
