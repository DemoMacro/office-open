import { DocParseError } from "../errors";
import type { NumberPair } from "../streams/pieces";
import { readUint32, requireRange } from "./binary";
export function parseCommentRanges(
  table: Uint8Array,
  range: NumberPair,
): Array<{ start: number; end: number; data: Uint8Array }> {
  const message = "Invalid DOC annotation references: they are outside the table stream";
  requireRange(table, range.offset, range.length, message);
  if (range.length < 8 || (range.length - 4) % 34 !== 0) {
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
  const count = (range.length - 4) / 34;
  const values: Array<{ start: number; end: number; data: Uint8Array }> = [];
  for (let index = 0; index < count; index += 1) {
    const start = readUint32(table, range.offset + index * 4, message);
    const end = readUint32(table, range.offset + (index + 1) * 4, message);
    const dataOffset = range.offset + (count + 1) * 4 + index * 30;
    values.push({
      start,
      end,
      data: table.subarray(dataOffset, dataOffset + 30),
    });
  }
  return values;
}
