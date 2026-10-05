import { DocParseError } from "../errors";
import type { NumberPair } from "../streams/pieces";
import { readUint32, requireRange } from "./binary";
export function parseCommentRanges(
  table: Uint8Array,
  range: NumberPair,
): Array<{ start: number; end: number; data: Uint8Array }> {
  const message = "Invalid DOC annotation references: they are outside the table stream";
  requireRange(table, range.offset, range.length, message);
  if (range.length < 8 || range.length % 4 !== 0) {
    throw new DocParseError("Invalid DOC annotation references: malformed CP count", {
      part: "table",
      path: "PlcfandRef",
      recordName: "PLC",
      offset: range.offset,
      length: range.length,
      byteRange: [range.offset, range.offset + range.length],
      reason: "invalid-record-length",
    });
  }
  const count = (range.length - 4) / 4;
  const descriptorLength = range.length - (count + 1) * 4;
  if (descriptorLength !== 0 && descriptorLength !== count * 4) {
    throw new DocParseError("Invalid DOC annotation references: truncated descriptors", {
      part: "table",
      path: "PlcfandRef",
      recordName: "FRD",
      offset: range.offset + (count + 1) * 4,
      length: count * 4,
      byteRange: [
        range.offset + (count + 1) * 4,
        range.offset + (count + 1) * 4 + descriptorLength,
      ],
      reason: "invalid-record-length",
    });
  }
  const values: Array<{ start: number; end: number; data: Uint8Array }> = [];
  for (let index = 0; index < count; index += 1) {
    const start = readUint32(table, range.offset + index * 4, message);
    const end = readUint32(table, range.offset + (index + 1) * 4, message);
    const dataOffset = range.offset + (count + 1) * 4 + index * 4;
    values.push({
      start,
      end,
      data: descriptorLength === 0 ? new Uint8Array(4) : table.subarray(dataOffset, dataOffset + 4),
    });
  }
  return values;
}
