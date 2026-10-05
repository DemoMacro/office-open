import { DocParseError } from "../errors";
import type { NumberPair } from "../streams/pieces";
import { readUint16, readUint32, requireRange } from "./binary";
import { FKP_LENGTH } from "./constants";
import type { BinTableEntry, TextboxBoundary } from "./models";
export function parseBinTable(
  table: Uint8Array,
  range: NumberPair,
  kind: "character" | "paragraph",
): BinTableEntry[] {
  const message = `Invalid DOC ${kind} bin table: it is outside the table stream`;
  requireRange(table, range.offset, range.length, message);
  if (range.length === 0) return [];
  if (range.length < 12 || (range.length - 4) % 8 !== 0) {
    throw new DocParseError(`Invalid DOC ${kind} bin table: malformed FC and page counts`, {
      part: "table",
      path: `Plcfbte${kind === "character" ? "Chpx" : "Papx"}`,
      offset: range.offset,
      length: range.length,
      byteRange: [range.offset, range.offset + range.length],
      reason: "invalid-record-length",
    });
  }
  const count = (range.length - 4) / 8;
  const entries: BinTableEntry[] = [];
  for (let index = 0; index < count; index++) {
    const fcStart = readUint32(table, range.offset + index * 4, message);
    const fcEnd = readUint32(table, range.offset + (index + 1) * 4, message);
    const page = readUint32(table, range.offset + (count + 1) * 4 + index * 4, message) & 0x3fffff;
    if (index > 0 && fcStart < entries[index - 1]!.fcEnd) {
      throw new DocParseError(`Invalid DOC ${kind} bin table: FC boundaries overlap`);
    }
    if (fcEnd <= fcStart || page > Number.MAX_SAFE_INTEGER / FKP_LENGTH) {
      throw new DocParseError(`Invalid DOC ${kind} bin table: invalid FC or page boundary`);
    }
    entries.push({ fcStart, fcEnd, page: page * FKP_LENGTH });
  }
  return entries;
}

export function parseNoteBoundaries(
  table: Uint8Array,
  range: NumberPair,
  storyLength: number,
): TextboxBoundary[] {
  if (storyLength === 0) return [];
  if (range.length === 0) {
    throw new DocParseError("Invalid Word document: note boundary table is missing");
  }
  const message = "Invalid DOC note table: it is outside the table stream";
  requireRange(table, range.offset, range.length, message);
  if (range.length < 12 || range.length % 4 !== 0) {
    throw new DocParseError("Invalid DOC note table: malformed CP count");
  }
  const cps = Array.from({ length: range.length / 4 }, (_, index) =>
    readUint32(table, range.offset + index * 4, message),
  );
  if (cps.slice(0, -2).some((cp) => cp >= storyLength)) {
    throw new DocParseError("Invalid DOC note table: CP boundary is outside the story");
  }
  if (cps[cps.length - 2] !== storyLength - 1) {
    throw new DocParseError("Invalid DOC note table: final boundary is invalid");
  }
  return cps.slice(0, -2).map((start, index) => ({
    start,
    end: cps[index + 1]!,
  }));
}

export function parseHeaderStreams(
  table: Uint8Array,
  plcfhddText: NumberPair,
  storyLength: number,
): TextboxBoundary[] {
  if (storyLength === 0) return [];
  if (plcfhddText.length === 0) {
    throw new DocParseError("Invalid Word document: header boundary table is missing");
  }
  const message = "Invalid DOC header table: it is outside the table stream";
  requireRange(table, plcfhddText.offset, plcfhddText.length, message);
  if (plcfhddText.length < 12 || plcfhddText.length % 4 !== 0) {
    throw new DocParseError("Invalid DOC header table: malformed CP count");
  }
  const cps = Array.from({ length: plcfhddText.length / 4 }, (_, index) =>
    readUint32(table, plcfhddText.offset + index * 4, message),
  );
  if (cps.slice(0, -2).some((cp) => cp >= storyLength)) {
    throw new DocParseError("Invalid DOC header table: CP boundary is outside the story");
  }
  if (cps[cps.length - 2] !== storyLength - 1) {
    throw new DocParseError("Invalid DOC header table: final boundary is invalid");
  }
  return cps.slice(0, -2).map((start, index) => ({
    start,
    end: cps[index + 1]!,
  }));
}

export function parseTextboxBoundaries(
  table: Uint8Array,
  range: NumberPair,
  storyLength: number,
): TextboxBoundary[] {
  if (range.length === 0) {
    throw new DocParseError("Invalid Word document: textbox boundary table is missing");
  }
  const message = "Invalid DOC textbox table: it is outside the table stream";
  requireRange(table, range.offset, range.length, message);
  if (range.length < 4 || (range.length - 4) % 26 !== 0) {
    throw new DocParseError("Invalid DOC textbox table: malformed CP and textbox counts", {
      part: "table",
      path: "PlcftxbxTxt",
      offset: range.offset,
      length: range.length,
      byteRange: [range.offset, range.offset + range.length],
      reason: "invalid-record-length",
    });
  }
  const count = (range.length - 4) / 26;
  const boundaries: TextboxBoundary[] = [];
  const cps: number[] = [];
  for (let index = 0; index <= count; index++) {
    const cp = readUint32(table, range.offset + index * 4, message);
    if (index > 0 && cp <= cps[index - 1]!) {
      throw new DocParseError("Invalid DOC textbox table: CP boundaries are not increasing");
    }
    cps.push(cp);
  }
  if (count === 0) return [];
  if (cps[0] !== 0 || cps[count]! > storyLength) {
    throw new DocParseError("Invalid DOC textbox table: boundaries do not fit the story");
  }
  const recordsOffset = range.offset + (count + 1) * 4;
  for (let index = 0; index < count - 1; index++) {
    const reusable =
      index === count - 1 || readUint16(table, recordsOffset + index * 22 + 8, message) !== 0;
    if (!reusable) boundaries.push({ start: cps[index]!, end: cps[index + 1]! });
  }
  return boundaries.filter((item) => item.end > item.start);
}
