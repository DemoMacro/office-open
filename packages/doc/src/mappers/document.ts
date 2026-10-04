import { CompoundFileReader } from "@office-open/core";
import type {
  DocumentOptions,
  ParagraphChild,
  ParagraphOptions,
  RunOptions,
  SectionChild,
  TableCellOptions,
  TableOptions,
  TableRowOptions,
} from "@office-open/docx";

import { readSummaryInformation } from "../cfb/container";
import { decryptWordStreams } from "../cfb/encryption";
import { DocParseError } from "../errors";
import type { LegacyDocumentOptions, LegacyRevisionRange } from "../records/models";
import {
  parseBookmarks,
  parseFields,
  parseLists,
  parsePictures,
  parseSectionProperties,
  parseStylesheet,
  type SectionModel,
  type StylesheetModel,
} from "../records/structures";

interface Fib {
  nFib: number;
  ccpText: number;
  ccpFootnotes: number;
  ccpHeaders: number;
  ccpComments: number;
  ccpEndnotes: number;
  ccpTextboxes: number;
  ccpHeaderTextboxes: number;
  totalCharacters: number;
  clx: NumberPair;
  characterBinTable: NumberPair;
  paragraphBinTable: NumberPair;
  footnoteTable: NumberPair;
  headerTable: NumberPair;
  endnoteTable: NumberPair;
  textboxTable: NumberPair;
  headerTextboxTable: NumberPair;
  styleSheet: NumberPair;
  sectionTable: NumberPair;
  bookmarkNames: NumberPair;
  bookmarkStarts: NumberPair;
  bookmarkEnds: NumberPair;
  fields: NumberPair;
  comments: NumberPair;
  drawing: NumberPair;
  list: NumberPair;
  listOverrides: NumberPair;
}

interface BinTableEntry {
  fcStart: number;
  fcEnd: number;
  page: number;
}

interface CharacterRange {
  cpStart: number;
  cpEnd: number;
  properties: RunOptions;
}

interface TextboxBoundary {
  start: number;
  end: number;
}

interface ParagraphRange {
  cpStart: number;
  cpEnd: number;
  properties: LegacyParagraphProperties;
}

interface LegacyParagraphProperties {
  alignment?: "left" | "center" | "right" | "both";
  indentLeft?: number;
  indentRight?: number;
  firstLine?: number;
  spacingBefore?: number;
  spacingAfter?: number;
  spacingLine?: number;
  outlineLevel?: number;
  indent?: {
    left?: number;
    right?: number;
    firstLine?: number;
  };
  spacing?: {
    before?: number;
    after?: number;
    line?: number;
    lineRule?: "auto";
  };
  inTable?: boolean;
  rowEnd?: boolean;
  columnWidths?: number[];
  cellMerges?: Array<{ restart?: boolean; continue?: boolean }>;
}

interface TextCharacter {
  value: string;
  cp: number;
  paragraphEnd?: boolean;
  hyperlink?: { url?: string; anchor?: string };
}

const FLAG_ENCRYPTED = 0x0100;
const FLAG_TABLE_ONE = 0x0200;
const FIB_SIGNATURE = 0xa5ec;
const FIB_BASE_LENGTH = 32;
const FKP_LENGTH = 512;
const COLOR_INDEXES = [
  "auto",
  "000000",
  "0000ff",
  "00ffff",
  "00ff00",
  "ff00ff",
  "ff0000",
  "ffff00",
  "ffffff",
  "000080",
  "008080",
  "008000",
  "800080",
  "800000",
  "808000",
  "808080",
  "c0c0c0",
] as const;

const WINDOWS_1252_HIGH = new Map<number, string>([
  [0x80, "\u20ac"],
  [0x82, "\u201a"],
  [0x83, "\u0192"],
  [0x84, "\u201e"],
  [0x85, "\u2026"],
  [0x86, "\u2020"],
  [0x87, "\u2021"],
  [0x88, "\u02c6"],
  [0x89, "\u2030"],
  [0x8a, "\u0160"],
  [0x8b, "\u2039"],
  [0x8c, "\u0152"],
  [0x8e, "\u017d"],
  [0x91, "\u2018"],
  [0x92, "\u2019"],
  [0x93, "\u201c"],
  [0x94, "\u201d"],
  [0x95, "\u2022"],
  [0x96, "\u2013"],
  [0x97, "\u2014"],
  [0x98, "\u02dc"],
  [0x99, "\u2122"],
  [0x9a, "\u0161"],
  [0x9b, "\u203a"],
  [0x9c, "\u0153"],
  [0x9e, "\u017e"],
  [0x9f, "\u0178"],
]);

function requireRange(bytes: Uint8Array, offset: number, length: number, message: string): void {
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

function readUint16(bytes: Uint8Array, offset: number, message: string): number {
  requireRange(bytes, offset, 2, message);
  return bytes[offset]! | (bytes[offset + 1]! << 8);
}

function readUint32(bytes: Uint8Array, offset: number, message: string): number {
  requireRange(bytes, offset, 4, message);
  return (
    (bytes[offset]! |
      (bytes[offset + 1]! << 8) |
      (bytes[offset + 2]! << 16) |
      (bytes[offset + 3]! << 24)) >>>
    0
  );
}

function parseFib(word: Uint8Array): Fib {
  const streamMessage = "Invalid Word document: FIB is truncated";
  requireRange(word, 0, FIB_BASE_LENGTH, streamMessage);
  if (readUint16(word, 0, streamMessage) !== FIB_SIGNATURE) {
    throw new DocParseError("Invalid Word document: FIB signature is not 0xA5EC", {
      part: "stream",
      path: "WordDocument",
      recordName: "FIB",
      offset: 0,
      length: 2,
      reason: "unsupported-required-structure",
    });
  }
  const flags = readUint16(word, 10, streamMessage);
  if ((flags & FLAG_ENCRYPTED) !== 0) {
    throw new DocParseError("Encrypted Word documents are not supported");
  }

  const csw = readUint16(word, 32, streamMessage);
  const cslwOffset = FIB_BASE_LENGTH + 2 + csw * 2;
  const cslw = readUint16(word, cslwOffset, streamMessage);
  const characterLengthsOffset = cslwOffset + 2;
  const blobCountOffset = characterLengthsOffset + cslw * 4;
  const cbRgFcLcb = readUint16(word, blobCountOffset, streamMessage);
  const blobOffset = blobCountOffset + 2;
  const readPair = (index: number, name: string): NumberPair => {
    if (index + 1 > cbRgFcLcb) {
      return { offset: 0, length: 0 };
    }
    const pairOffset = blobOffset + index * 8;
    const message = `Invalid Word document: ${name} is outside the WordDocument stream`;
    return {
      offset: readUint32(word, pairOffset, message),
      length: readUint32(word, pairOffset + 4, message),
    };
  };

  if (cslw < 11) {
    throw new DocParseError("Invalid Word document: FIB text lengths are truncated");
  }
  const readLength = (index: number): number =>
    readUint32(word, characterLengthsOffset + index * 4, streamMessage);
  const ccpText = readLength(3);
  const ccpFootnotes = readLength(4);
  const ccpHeaders = readLength(5);
  const ccpComments = readLength(7);
  const ccpEndnotes = readLength(8);
  const ccpTextboxes = readLength(9);
  const ccpHeaderTextboxes = readLength(10);
  return {
    ccpText,
    ccpFootnotes,
    ccpHeaders,
    ccpComments,
    ccpEndnotes,
    ccpTextboxes,
    ccpHeaderTextboxes,
    totalCharacters:
      ccpText +
      ccpFootnotes +
      ccpHeaders +
      ccpComments +
      ccpEndnotes +
      ccpTextboxes +
      ccpHeaderTextboxes,
    nFib: readUint16(word, 2, streamMessage),
    characterBinTable: readPair(12, "fcPlcfbteChpx"),
    paragraphBinTable: readPair(13, "fcPlcfbtePapx"),
    footnoteTable: readPair(3, "fcPlcffndTxt"),
    headerTable: readPair(11, "fcPlcfHdd"),
    endnoteTable: readPair(47, "fcPlcfendTxt"),
    clx: readPair(33, "fcClx"),
    textboxTable: readPair(56, "fcPlcftxbxTxt"),
    headerTextboxTable: readPair(58, "fcPlcfHdrTxbxTxt"),
    styleSheet: readPair(1, "fcStshf"),
    sectionTable: readPair(6, "fcPlcfSed"),
    bookmarkNames: readPair(21, "fcSttbfBkmk"),
    bookmarkStarts: readPair(22, "fcPlcfBkf"),
    bookmarkEnds: readPair(23, "fcPlcfBkl"),
    fields: readPair(16, "fcPlcffldMom"),
    comments: readPair(4, "fcPlcfandRef"),
    drawing: readPair(50, "fcDggInfo"),
    list: readPair(73, "fcPlfLst"),
    listOverrides: readPair(74, "fcPlcflfo"),
  };
}

interface LegacyFib {
  fcMin: number;
  fcMac: number;
  ccpText: number;
}

function parseLegacyFib(word: Uint8Array): LegacyFib {
  const message = "Invalid legacy Word document: FIB is truncated";
  requireRange(word, 0, 88, message);
  const nFib = readUint16(word, 2, message);
  if (
    readUint16(word, 0, message) !== 0xa5dc ||
    nFib < 101 ||
    nFib > 105 ||
    (readUint16(word, 10, message) & FLAG_ENCRYPTED) !== 0
  ) {
    throw new DocParseError("Invalid Word document: unsupported FIB signature", {
      part: "stream",
      path: "WordDocument",
      recordName: "FIB",
      offset: 0,
      length: 2,
      reason: "unsupported-required-structure",
    });
  }
  return {
    fcMin: readUint32(word, 24, message),
    fcMac: readUint32(word, 28, message),
    ccpText: readUint32(word, 52, message),
  };
}

function parseLegacyText(word: Uint8Array, fib: LegacyFib): string {
  requireRange(
    word,
    fib.fcMin,
    fib.fcMac - fib.fcMin,
    "Invalid legacy Word document: text range is outside the stream",
  );
  let result = "";
  for (let index = fib.fcMin; index < fib.fcMac; index += 1) {
    const byte = word[index]!;
    result += WINDOWS_1252_HIGH.get(byte) ?? String.fromCharCode(byte);
  }
  return result;
}

function legacyChildren(text: string): SectionChild[] {
  const children: SectionChild[] = [];
  for (const value of text.split(/\r/)) {
    children.push({ paragraph: value });
  }
  return children;
}

import { parsePieceTable, type NumberPair, type Piece } from "../streams/pieces";

function parseBinTable(
  table: Uint8Array,
  range: NumberPair,
  kind: "character" | "paragraph",
): BinTableEntry[] {
  const message = `Invalid DOC ${kind} bin table: it is outside the table stream`;
  requireRange(table, range.offset, range.length, message);
  if (range.length === 0) return [];
  if (range.length < 12 || (range.length - 4) % 8 !== 0) {
    throw new DocParseError(`Invalid DOC ${kind} bin table: malformed FC and page counts`);
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

function parseNoteBoundaries(
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

function parseHeaderStreams(
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

function parseTextboxBoundaries(
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
    throw new DocParseError("Invalid DOC textbox table: malformed CP and textbox counts");
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

function fcToCp(fc: number, piece: Piece): number {
  const characters = piece.compressed ? fc - piece.fc : (fc - piece.fc) / 2;
  if (!Number.isInteger(characters)) {
    throw new DocParseError("Invalid DOC property range: FC is not aligned to the text piece");
  }
  return piece.cpStart + characters;
}

function overlappingPieces(
  pieces: Piece[],
  cpEndLimit: number,
  fcStart: number,
  fcEnd: number,
): Array<{ cpStart: number; cpEnd: number; piece: Piece }> {
  const result: Array<{ cpStart: number; cpEnd: number; piece: Piece }> = [];
  for (const piece of pieces) {
    const cpEnd = Math.min(piece.cpEnd, cpEndLimit);
    if (piece.cpStart >= cpEnd) continue;
    const pieceFcEnd = piece.fc + (cpEnd - piece.cpStart) * (piece.compressed ? 1 : 2);
    if (piece.fc >= fcEnd || pieceFcEnd <= fcStart) continue;
    const cpIntersectionStart = Math.max(piece.cpStart, fcToCp(Math.max(fcStart, piece.fc), piece));
    const cpIntersectionEnd = Math.min(cpEnd, fcToCp(Math.min(fcEnd, pieceFcEnd), piece));
    if (cpIntersectionEnd > cpIntersectionStart) {
      result.push({ cpStart: cpIntersectionStart, cpEnd: cpIntersectionEnd, piece });
    }
  }
  return result;
}

function parseCharacterFkp(
  word: Uint8Array,
  page: number,
  pieces: Piece[],
  cpEndLimit: number,
): CharacterRange[] {
  requireRange(
    word,
    page,
    FKP_LENGTH,
    "Invalid DOC character FKP: page is outside the WordDocument stream",
  );
  const count = word[page + FKP_LENGTH - 1]!;
  if (count === 0 || count > (FKP_LENGTH - 5) / 5) {
    throw new DocParseError("Invalid DOC character FKP: invalid entry count");
  }
  const ranges: CharacterRange[] = [];
  for (let index = 0; index < count; index++) {
    const fcStart = readUint32(
      word,
      page + index * 4,
      "Invalid DOC character FKP: FC array is truncated",
    );
    const fcEnd = readUint32(
      word,
      page + (index + 1) * 4,
      "Invalid DOC character FKP: FC array is truncated",
    );
    const propertyOffset = word[page + (count + 1) * 4 + index]!;
    if (fcEnd <= fcStart) throw new DocParseError("Invalid DOC character FKP: invalid FC range");
    if (propertyOffset === 0) continue;
    const propertiesStart = page + propertyOffset * 2;
    const propertiesLength = word[propertiesStart]!;
    const grpprlStart = propertiesStart + 1;
    requireRange(
      word,
      grpprlStart,
      propertiesLength,
      "Invalid DOC character FKP: property data is outside the page",
    );
    const properties = parseCharacterProperties(
      word.subarray(grpprlStart, grpprlStart + propertiesLength),
    );
    for (const overlap of overlappingPieces(pieces, cpEndLimit, fcStart, fcEnd)) {
      ranges.push({ ...overlap, properties });
    }
  }
  return ranges.sort((left, right) => left.cpStart - right.cpStart || left.cpEnd - right.cpEnd);
}

function parseParagraphFkp(word: Uint8Array, page: number): ParagraphRange[] {
  requireRange(
    word,
    page,
    FKP_LENGTH,
    "Invalid DOC paragraph FKP: page is outside the WordDocument stream",
  );
  const count = word[page + FKP_LENGTH - 1]!;
  if (count === 0 || count > Math.floor((FKP_LENGTH - 5) / 17)) {
    throw new DocParseError("Invalid DOC paragraph FKP: invalid entry count");
  }
  const ranges: ParagraphRange[] = [];
  for (let index = 0; index < count; index++) {
    const fcStart = readUint32(
      word,
      page + index * 4,
      "Invalid DOC paragraph FKP: FC array is truncated",
    );
    const fcEnd = readUint32(
      word,
      page + (index + 1) * 4,
      "Invalid DOC paragraph FKP: FC array is truncated",
    );
    const entryOffset = page + (count + 1) * 4 + index * 13;
    const propertyOffset = word[entryOffset]!;
    if (fcEnd <= fcStart) throw new DocParseError("Invalid DOC paragraph FKP: invalid FC range");
    if (propertyOffset === 0) {
      ranges.push({ cpStart: fcStart, cpEnd: fcEnd, properties: {} });
      continue;
    }
    const propertiesStart = page + propertyOffset * 2;
    const wordCount = word[propertiesStart]!;
    const grpprlStart = propertiesStart + (wordCount === 0 ? 4 : 3);
    const propertiesLength =
      wordCount === 0 ? word[propertiesStart + 1]! * 2 - 2 : wordCount * 2 - 3;
    requireRange(
      word,
      grpprlStart,
      propertiesLength,
      "Invalid DOC paragraph FKP: property data is outside the page",
    );
    const properties = parseParagraphProperties(
      word.subarray(grpprlStart, grpprlStart + propertiesLength),
    );
    ranges.push({ cpStart: fcStart, cpEnd: fcEnd, properties });
  }
  return ranges;
}

const ALIGNMENT_VALUES = ["left", "center", "right", "both"] as const;

function parseParagraphProperties(papxText: Uint8Array): LegacyParagraphProperties {
  const properties: LegacyParagraphProperties = {};
  iterateProperties(papxText, (opcode, operandOffset, length) => {
    const operation = opcode & 0x00ff;
    const signed = () =>
      (readUint16(papxText, operandOffset, "Invalid DOC paragraph properties") << 16) >> 16;
    if (operation === 0x03 && length === 1) {
      properties.alignment = ALIGNMENT_VALUES[papxText[operandOffset]! & 3];
    } else if (operation === 0x0f || operation === 0x5e) {
      properties.indentLeft = signed();
    } else if (operation === 0x0e || operation === 0x5d) {
      properties.indentRight = signed();
    } else if (operation === 0x11 || operation === 0x60) {
      properties.firstLine = Math.max(0, signed());
    } else if (operation === 0x13 && length === 2) {
      properties.spacingBefore = signed();
    } else if (operation === 0x14 && length === 2) {
      properties.spacingAfter = signed();
    } else if (operation === 0x12 && length >= 4) {
      properties.spacingLine = readUint16(
        papxText,
        operandOffset + 2,
        "Invalid DOC paragraph properties",
      );
    } else if (operation === 0x40 && length === 1) {
      properties.outlineLevel = papxText[operandOffset]! & 15;
    } else if (operation === 0x16 && length === 1) {
      properties.inTable = papxText[operandOffset] !== 0;
    } else if (operation === 0x17 && length === 1) {
      properties.rowEnd = papxText[operandOffset] !== 0;
    } else if (operation === 0x08 && length > 2) {
      parseTableProperties(
        papxText.subarray(operandOffset + 1, operandOffset + length),
        properties,
      );
    }
  });
  return properties;
}

function parseTableProperties(tap: Uint8Array, properties: LegacyParagraphProperties): void {
  if (tap.byteLength < 4) return;
  const cellCount = tap[0]!;
  if (cellCount === 0 || 2 + (cellCount + 1) * 2 > tap.byteLength) return;
  properties.columnWidths = [];
  let previous = tap[2]! | (tap[3]! << 8);
  for (let index = 0; index < cellCount; index += 1) {
    const offset = 2 + (index + 1) * 2;
    const boundary = tap[offset]! | (tap[offset + 1]! << 8);
    properties.columnWidths.push(Math.max(0, boundary - previous));
    previous = boundary;
  }
  properties.inTable = true;
  const tcOffset = 2 + (cellCount + 1) * 2;
  properties.cellMerges = [];
  for (let index = 0; tcOffset + index + 1 < tap.byteLength && index < cellCount; index += 1) {
    const flags = tap[tcOffset + index]! | (tap[tcOffset + index + 1]! << 8);
    properties.cellMerges.push({
      restart: (flags & 0x0002) !== 0,
      continue: (flags & 0x0001) !== 0,
    });
  }
}

function operandLength(opcode: number, bytes: Uint8Array, offset: number): number {
  switch ((opcode >> 13) & 0b111) {
    case 0:
    case 1:
      return 1;
    case 2:
    case 4:
    case 5:
      return 2;
    case 3:
      return 4;
    case 7:
      return 3;
    case 6:
      requireRange(bytes, offset, 1, "Invalid DOC properties: truncated variable operand");
      return 1 + bytes[offset]!;
    default:
      throw new DocParseError("Invalid DOC properties: unknown operand form");
  }
}

function iterateProperties(
  grpprl: Uint8Array,
  visit?: (opcode: number, operandOffset: number, operandLength: number) => void,
): void {
  let offset = 0;
  while (offset < grpprl.byteLength) {
    if (grpprl.byteLength - offset < 2) {
      if (grpprl.byteLength - offset === 1) break;
      throw new DocParseError("Invalid DOC properties: truncated property opcode");
    }
    const opcode = readUint16(grpprl, offset, "Invalid DOC properties: truncated property opcode");
    const length = operandLength(opcode, grpprl, offset + 2);
    const operandOffset = offset + 2;
    requireRange(
      grpprl,
      operandOffset,
      length,
      "Invalid DOC properties: truncated property operand",
    );
    visit?.(opcode, operandOffset, length);
    offset = operandOffset + length;
  }
}

function parseCharacterProperties(chpxText: Uint8Array): RunOptions {
  const properties: RunOptions = {};
  iterateProperties(chpxText, (opcode, operandOffset, length) => {
    if (opcode === 0x0835 && length === 1)
      properties.bold = chpxText[operandOffset]! === 1 || chpxText[operandOffset]! === 128;
    if (opcode === 0x0836 && length === 1)
      properties.italic = chpxText[operandOffset]! === 1 || chpxText[operandOffset]! === 128;
    if (opcode === 0x0837 && length === 1)
      properties.strike = chpxText[operandOffset]! === 1 || chpxText[operandOffset]! === 128;
    if (opcode === 0x4a43 && length === 2) {
      properties.size = readUint16(chpxText, operandOffset, "Invalid DOC character properties") / 2;
    }
    if (opcode === 0x6a03 && length === 2) {
      const colorIndex = chpxText[operandOffset]!;
      if (colorIndex < COLOR_INDEXES.length) properties.color = COLOR_INDEXES[colorIndex]!;
    }
    if (opcode === 0x2a47 && length === 1) {
      const style = chpxText[operandOffset]!;
      properties.underline =
        style === 0 || style === undefined
          ? undefined
          : { type: style === 2 ? "words" : style === 3 ? "double" : "single" };
    }
    if (opcode === 0x2a48 && length === 1) {
      properties.highlight = chpxText[operandOffset] === 0 ? undefined : "yellow";
    }
    if (opcode === 0x4a4f && length === 2) {
      properties.font = `font-${readUint16(chpxText, operandOffset, "Invalid DOC character properties")}`;
    }
  });
  return properties;
}

function decodePiece(word: Uint8Array, piece: Piece): string {
  const byteLength = (piece.cpEnd - piece.cpStart) * (piece.compressed ? 1 : 2);
  requireRange(
    word,
    piece.fc,
    byteLength,
    "Invalid DOC piece: text is outside the WordDocument stream",
  );
  let result = "";
  if (piece.compressed) {
    for (let index = 0; index < byteLength; index++) {
      const byte = word[piece.fc + index]!;
      result += WINDOWS_1252_HIGH.get(byte) ?? String.fromCharCode(byte);
    }
    return result;
  }
  for (let index = 0; index < byteLength; index += 2) {
    result += String.fromCharCode(word[piece.fc + index]! | (word[piece.fc + index + 1]! << 8));
  }
  return result;
}

function normalizeCharacter(value: string): string | undefined {
  switch (value) {
    case "\r":
    case "\x07":
    case "\x0c":
      return value;
    case "\x1e":
      return "\u2011";
    case "\x1f":
    case "\x01":
    case "\x02":
    case "\x05":
    case "\x08":
    case "\x13":
    case "\x14":
    case "\x15":
      return undefined;
    default:
      return value >= " " ? value : undefined;
  }
}

function projectText(pieces: Piece[], word: Uint8Array, ccpText: number): TextCharacter[] {
  const characters: TextCharacter[] = [];
  const fieldResults: boolean[] = [];
  for (const piece of pieces) {
    if (piece.cpStart >= ccpText) break;
    const cpEnd = Math.min(piece.cpEnd, ccpText);
    const decoded = decodePiece(word, { ...piece, cpEnd });
    for (let index = 0; index < decoded.length; index++) {
      const source = decoded[index]!;
      if (source === "\x13") {
        fieldResults.push(false);
        continue;
      }
      if (source === "\x14") {
        if (fieldResults.length > 0) fieldResults[fieldResults.length - 1] = true;
        continue;
      }
      if (source === "\x15") {
        fieldResults.pop();
        continue;
      }
      if (fieldResults.at(-1) === false) continue;
      const value = normalizeCharacter(source);
      const isParagraphEnd = value === "\r" || value === "\x07" || value === "\x0c";
      if (isParagraphEnd) {
        characters.push({ value: "", cp: piece.cpStart + index, paragraphEnd: true });
      } else if (value !== undefined) {
        characters.push({ value, cp: piece.cpStart + index });
      }
    }
  }
  return characters;
}

function projectChildren(
  characters: TextCharacter[],
  ranges: readonly CharacterRange[],
  paragraphRanges: readonly ParagraphRange[] = [],
): SectionChild[] {
  const children: SectionChild[] = [];
  let paragraph: TextCharacter[] = [];

  const flushParagraph = (): void => {
    if (paragraph.length === 0) {
      children.push({
        paragraph:
          paragraphRanges.length > 0
            ? { children: [], ...paragraphPropertiesFor(paragraphRanges, paragraph[0]?.cp) }
            : "",
      });
      return;
    }
    const boundaries = new Set<number>([0, paragraph.length]);
    for (let index = 0; index < paragraph.length; index++) {
      const cp = paragraph[index]!.cp;
      for (const range of ranges) {
        if (cp === range.cpStart) boundaries.add(index);
      }
      if (index > 0 && paragraph[index]!.hyperlink !== paragraph[index - 1]!.hyperlink) {
        boundaries.add(index);
      }
    }
    const sortedBoundaries = [...boundaries].sort((left, right) => left - right);
    const runs: ParagraphChild[] = [];
    for (let index = 0; index < sortedBoundaries.length - 1; index++) {
      const start = sortedBoundaries[index]!;
      const end = sortedBoundaries[index + 1]!;
      const text = paragraph
        .slice(start, end)
        .map((character) => character.value)
        .join("");
      if (text.length === 0) continue;
      const cp = paragraph[start]!.cp;
      const range = ranges.find((candidate) => candidate.cpStart <= cp && cp < candidate.cpEnd);
      const hyperlink = paragraph[start]!.hyperlink;
      const breakCount = text.split("\n").length - 1;
      const visibleText = text.replaceAll("\n", "");
      if (breakCount > 0) runs.push({ break: breakCount });
      if (visibleText.length > 0 && hyperlink) {
        runs.push({
          hyperlink: { ...hyperlink },
          children: [{ text: visibleText, ...range?.properties }],
        });
      } else if (visibleText.length > 0) {
        runs.push({ text: visibleText, ...range?.properties });
      }
    }
    const plainText = paragraph.map((character) => character.value).join("");
    const hasProperties = runs.some((run) =>
      Object.keys(run).some((key) => key !== "text" && key !== "break"),
    );
    const paragraphProperties = paragraphPropertiesFor(paragraphRanges, paragraph[0]?.cp);
    children.push({
      paragraph:
        hasProperties || runs.length > 1 || Object.keys(paragraphProperties).length > 0
          ? { children: runs, ...paragraphProperties }
          : plainText,
    });
    paragraph = [];
  };

  for (const character of characters) {
    if (character.paragraphEnd) flushParagraph();
    else paragraph.push(character);
  }
  if (paragraph.length > 0) flushParagraph();
  if (children.length === 0) children.push({ paragraph: "" });
  return children;
}

function projectStory(
  characters: readonly TextCharacter[],
  ranges: readonly CharacterRange[],
  start: number,
  end: number,
  paragraphRanges: readonly ParagraphRange[] = [],
): SectionChild[] {
  return projectChildren(
    characters.filter((character) => character.cp >= start && character.cp < end),
    ranges,
    paragraphRanges,
  );
}

function applyHyperlinkFields(
  characters: TextCharacter[],
  fields: readonly { start: number; end: number; instruction: string }[],
  rawCharacters: readonly string[],
): void {
  for (const field of fields) {
    const match = /^HYPERLINK\s+(?:"([^"]*)"|(\S+))/i.exec(field.instruction.trim());
    if (!match) continue;
    const target = match[2] ?? match[1];
    if (!target) continue;
    const hyperlink = target.startsWith("#") ? { anchor: target.slice(1) } : { url: target };
    for (
      let cp = field.end + 1;
      cp < rawCharacters.length && rawCharacters[cp] !== "\x15";
      cp += 1
    ) {
      const character = characters.find((candidate) => candidate.cp === cp);
      if (character) character.hyperlink = hyperlink;
    }
  }
}

function parseOptional<T>(parse: () => T, fallback: T): T {
  try {
    return parse();
  } catch (error) {
    if (!(error instanceof DocParseError)) throw error;
    return fallback;
  }
}

function paragraphPropertiesFor(
  ranges: readonly ParagraphRange[],
  cp: number | undefined,
): LegacyParagraphProperties {
  if (cp === undefined) return {};
  const range = ranges.find((candidate) => candidate.cpStart <= cp && cp < candidate.cpEnd);
  if (!range) return {};
  const {
    alignment,
    indentLeft,
    indentRight,
    firstLine,
    spacingBefore,
    spacingAfter,
    spacingLine,
    outlineLevel,
  } = range.properties;
  const properties: LegacyParagraphProperties = {};
  if (alignment !== undefined) properties.alignment = alignment;
  if (indentLeft !== undefined || indentRight !== undefined || firstLine !== undefined) {
    properties.indent = {
      ...(indentLeft !== undefined ? { left: indentLeft } : {}),
      ...(indentRight !== undefined ? { right: indentRight } : {}),
      ...(firstLine !== undefined ? { firstLine } : {}),
    };
  }
  if (spacingBefore !== undefined || spacingAfter !== undefined || spacingLine !== undefined) {
    properties.spacing = {
      ...(spacingBefore !== undefined ? { before: spacingBefore } : {}),
      ...(spacingAfter !== undefined ? { after: spacingAfter } : {}),
      ...(spacingLine !== undefined ? { line: spacingLine, lineRule: "auto" as const } : {}),
    };
  }
  if (outlineLevel !== undefined) properties.outlineLevel = outlineLevel;
  return properties;
}

function applyTables(
  children: SectionChild[],
  paragraphRanges: readonly ParagraphRange[],
  characters: readonly TextCharacter[],
): SectionChild[] {
  const rowByStart = new Map<number, ParagraphRange>();
  for (const range of paragraphRanges) {
    if (
      !range.properties.inTable ||
      !range.properties.rowEnd ||
      !range.properties.columnWidths?.length
    )
      continue;
    const existing = rowByStart.get(range.cpStart);
    rowByStart.set(range.cpStart, !existing || range.cpEnd > existing.cpEnd ? range : existing);
  }
  const rows: ParagraphRange[] = [];
  for (const range of [...rowByStart.values()].sort(
    (left, right) => left.cpStart - right.cpStart,
  )) {
    const previous = rows.at(-1);
    if (previous && range.cpStart < previous.cpEnd) {
      previous.cpEnd = Math.max(previous.cpEnd, range.cpEnd);
      previous.properties = { ...previous.properties, ...range.properties };
    } else {
      rows.push(range);
    }
  }
  if (rows.length === 0) return children;
  const result = [...children];
  let offset = 0;
  let previousEnd = 0;
  for (const row of rows) {
    const startIndex = paragraphIndexForCp(characters, previousEnd);
    const endIndex = paragraphIndexForCp(characters, row.cpEnd);
    const content = result.slice(startIndex + offset, endIndex + offset + 1);
    const columnCount = row.properties.columnWidths!.length;
    const cells: TableCellOptions[] = Array.from({ length: columnCount }, (_, index) => {
      const merge = row.properties.cellMerges?.[index];
      const chunkStart = Math.ceil((index * content.length) / columnCount);
      const chunkEnd = Math.ceil(((index + 1) * content.length) / columnCount);
      return {
        children: content.slice(chunkStart, chunkEnd),
        ...(merge?.continue ? { verticalMerge: "continue" as const } : {}),
        ...(merge?.restart ? { verticalMerge: "restart" as const } : {}),
      };
    });
    const table: SectionChild = {
      table: {
        columnWidths: row.properties.columnWidths,
        rows: [{ cells }] satisfies TableRowOptions[],
      } satisfies TableOptions,
    };
    result.splice(startIndex + offset, content.length, table);
    offset += 1 - content.length;
    previousEnd = row.cpEnd;
  }
  return result;
}

function paragraphIndexForCp(characters: readonly TextCharacter[], cp: number): number {
  return characters.filter((character) => character.paragraphEnd && character.cp < cp).length;
}

function projectCommentStory(
  characters: readonly TextCharacter[],
  ranges: readonly CharacterRange[],
  start: number,
  end: number,
): Array<string | ParagraphOptions> {
  return projectStory(characters, ranges, start, end).flatMap((child) =>
    "paragraph" in child ? [child.paragraph] : [],
  );
}

export function parseInternal(data: Uint8Array, password?: string): DocumentOptions {
  let reader: CompoundFileReader;
  try {
    reader = new CompoundFileReader(data);
  } catch (error) {
    throw new DocParseError(
      `Input is not a supported Compound File Binary document: ${(error as Error).message}`,
      { part: "container", path: "/", reason: "invalid-container" },
    );
  }
  if (!reader.entry("WordDocument")) {
    throw new DocParseError("Invalid Word document: WordDocument stream is missing", {
      part: "stream",
      path: "WordDocument",
      reason: "missing-required-stream",
    });
  }
  let word = reader.read("WordDocument");
  const legacy = readUint16(word, 0, "Invalid Word document: FIB is truncated") === 0xa5dc;
  if (legacy) {
    const fib = parseLegacyFib(word);
    return { sections: [{ children: legacyChildren(parseLegacyText(word, fib)) }] };
  }
  const flags = readUint16(word, 10, "Invalid Word document: FIB is truncated");
  let tablePath = (flags & FLAG_TABLE_ONE) !== 0 ? "1Table" : "0Table";
  let table: Uint8Array | undefined = reader.entry(tablePath) ? reader.read(tablePath) : undefined;
  let dataStream = reader.entry("Data") ? reader.read("Data") : undefined;
  if ((flags & FLAG_ENCRYPTED) !== 0) {
    if (!table) {
      throw new DocParseError("Encrypted Word documents are not supported", {
        part: "table",
        path: tablePath,
        reason: "encrypted-unsupported",
      });
    }
    const decrypted = decryptWordStreams(word, table, reader, password);
    word = decrypted.word;
    table = decrypted.table;
    tablePath = decrypted.tablePath;
    dataStream = decrypted.data;
  }
  const fib = parseFib(word);
  const metadata = readSummaryInformation(reader);
  if (!table) {
    throw new DocParseError(`Invalid Word document: ${tablePath} stream is missing`, {
      part: "table",
      path: tablePath,
      reason: "missing-required-stream",
    });
  }
  if (fib.totalCharacters === 0) return { sections: [{ children: [{ paragraph: "" }] }] };

  const pieces = parsePieceTable(table, fib.clx);
  if (fib.totalCharacters > pieces[pieces.length - 1]!.cpEnd) {
    throw new DocParseError("Invalid Word document: main text extends beyond the piece table");
  }
  const characters = projectText(pieces, word, fib.totalCharacters);
  const characterEntries = parseOptional(
    () => parseBinTable(table, fib.characterBinTable, "character"),
    [],
  );
  const characterRanges = characterEntries.flatMap((entry) =>
    parseOptional(() => parseCharacterFkp(word, entry.page, pieces, fib.totalCharacters), []),
  );
  const paragraphRanges = parseOptional(
    () => parseBinTable(table, fib.paragraphBinTable, "paragraph"),
    [],
  )
    .flatMap((entry) => parseOptional(() => parseParagraphFkp(word, entry.page), []))
    .map((range) => {
      const overlaps = overlappingPieces(pieces, fib.totalCharacters, range.cpStart, range.cpEnd);
      if (overlaps.length === 0) return undefined;
      return {
        cpStart: Math.min(...overlaps.map((overlap) => overlap.cpStart)),
        cpEnd: Math.max(...overlaps.map((overlap) => overlap.cpEnd)),
        properties: range.properties,
      };
    })
    .filter((range) => range !== undefined);

  const ranges = [
    { start: 0, end: fib.ccpText },
    { start: fib.ccpText, end: fib.ccpText + fib.ccpFootnotes },
    {
      start: fib.ccpText + fib.ccpFootnotes,
      end: fib.ccpText + fib.ccpFootnotes + fib.ccpHeaders,
    },
  ];
  const rawCharacters: string[] = [];
  for (const piece of pieces) {
    const cpEnd = Math.min(piece.cpEnd, fib.totalCharacters);
    const characters = decodePiece(word, { ...piece, cpEnd });
    for (let index = 0; index < characters.length; index += 1) {
      rawCharacters.push(characters[index]!);
    }
  }
  let sections: SectionModel[] = [];
  if (fib.sectionTable.length > 0) {
    try {
      sections = parseSectionProperties(table, fib.sectionTable, word).sections;
    } catch {
      sections = [];
    }
  }
  const sectionModels: SectionModel[] = sections.length > 0 ? sections : [{}];
  let stylesheet: StylesheetModel = { styles: [] };
  if (fib.styleSheet.length > 0) {
    try {
      stylesheet = parseStylesheet(table, fib.styleSheet);
    } catch {
      stylesheet = { styles: [] };
    }
  }
  let bookmarks: ReturnType<typeof parseBookmarks> = [];
  if (fib.bookmarkNames.length > 0) {
    try {
      bookmarks = parseBookmarks(table, fib.bookmarkNames, fib.bookmarkStarts, fib.bookmarkEnds);
    } catch {
      bookmarks = [];
    }
  }
  let fields: ReturnType<typeof parseFields> = [];
  if (fib.fields.length > 0) {
    try {
      fields = parseFields(rawCharacters, table, fib.fields);
    } catch {
      fields = [];
    }
  }
  applyHyperlinkFields(characters, fields, rawCharacters);
  let revisions: LegacyRevisionRange[] = [];
  if (fib.comments.length > 0) {
    try {
      revisions = parseCommentRanges(table, fib.comments).map((comment) => ({
        start: comment.start,
        end: comment.end,
        inserted: (comment.data[1]! & 0x01) === 0,
      }));
    } catch {
      revisions = [];
    }
  }
  const drawing = fib.drawing.length > 0 ? dataStream : undefined;
  const pictures =
    fib.drawing.length > 0 && drawing
      ? parsePictures(
          table.subarray(fib.drawing.offset, fib.drawing.offset + fib.drawing.length),
          drawing,
        )
      : [];
  let lists: ReturnType<typeof parseLists> | undefined;
  if (fib.list.length > 0 || fib.listOverrides.length > 0) {
    try {
      lists = parseLists(table, fib.list, fib.listOverrides);
    } catch {
      lists = undefined;
    }
  }
  const footnoteStart = ranges[1]!.start;
  const headerStart = ranges[2]!.start;
  const commentStart = headerStart + fib.ccpHeaders;
  const endnoteStart = commentStart + fib.ccpComments;
  const textboxStart = endnoteStart + fib.ccpEndnotes;
  const headerTextboxStart = textboxStart + fib.ccpTextboxes;

  const projectedChildren = projectStory(
    characters,
    characterRanges,
    0,
    fib.ccpText,
    paragraphRanges,
  );
  const sectionChildren = applyTables(projectedChildren, paragraphRanges, characters);
  const sectionChildrenWithBookmarks = applyBookmarks(sectionChildren, bookmarks, characters);
  const headerStories = parseOptional(
    () => parseHeaderStreams(table, fib.headerTable, fib.ccpHeaders),
    [],
  );
  const headerStoryChildren = (index: number): SectionChild[] => {
    const story = headerStories[index];
    if (!story) return [];
    return projectStory(
      characters,
      characterRanges,
      headerStart + story.start,
      headerStart + story.end,
    ).filter((child) => !("paragraph" in child && child.paragraph === ""));
  };
  const defaultHeaderChildren = headerStoryChildren(7);
  const evenHeaderChildren = headerStoryChildren(6);
  const firstHeaderChildren = headerStoryChildren(10);
  const defaultFooterChildren = headerStoryChildren(9);
  const evenFooterChildren = headerStoryChildren(8);
  const firstFooterChildren = headerStoryChildren(11);
  const textboxBoundaries =
    fib.ccpTextboxes > 0
      ? parseOptional(() => parseTextboxBoundaries(table, fib.textboxTable, fib.ccpTextboxes), [])
      : [];
  for (const boundary of textboxBoundaries) {
    sectionChildren.push({
      textbox: {
        children: projectStory(
          characters,
          characterRanges,
          textboxStart + boundary.start,
          textboxStart + boundary.end,
        ),
      },
    });
  }
  const headerTextboxBoundaries =
    fib.ccpHeaderTextboxes > 0
      ? parseOptional(
          () => parseTextboxBoundaries(table, fib.headerTextboxTable, fib.ccpHeaderTextboxes),
          [],
        )
      : [];
  for (const boundary of headerTextboxBoundaries) {
    defaultHeaderChildren.push({
      textbox: {
        children: projectStory(
          characters,
          characterRanges,
          headerTextboxStart + boundary.start,
          headerTextboxStart + boundary.end,
        ),
      },
    });
  }

  const sectionGeometry = sectionModels[0] ?? {};
  const hasPageSize = sectionGeometry.width !== undefined && sectionGeometry.height !== undefined;
  const hasPageMargin =
    sectionGeometry.marginTop !== undefined ||
    sectionGeometry.marginRight !== undefined ||
    sectionGeometry.marginBottom !== undefined ||
    sectionGeometry.marginLeft !== undefined;
  const sectionProperties = {
    ...(hasPageSize
      ? {
          pageSize: {
            width: sectionGeometry.width!,
            height: sectionGeometry.height!,
            ...(sectionGeometry.orientation ? { orientation: sectionGeometry.orientation } : {}),
          },
        }
      : {}),
    ...(hasPageMargin
      ? {
          pageMargin: {
            ...(sectionGeometry.marginTop !== undefined ? { top: sectionGeometry.marginTop } : {}),
            ...(sectionGeometry.marginRight !== undefined
              ? { right: sectionGeometry.marginRight }
              : {}),
            ...(sectionGeometry.marginBottom !== undefined
              ? { bottom: sectionGeometry.marginBottom }
              : {}),
            ...(sectionGeometry.marginLeft !== undefined
              ? { left: sectionGeometry.marginLeft }
              : {}),
          },
        }
      : {}),
  };
  const result: LegacyDocumentOptions = {
    sections: [
      {
        children: sectionChildrenWithBookmarks,
        ...(Object.keys(sectionProperties).length > 0 ? { properties: sectionProperties } : {}),
        ...(defaultHeaderChildren.length > 0 ||
        evenHeaderChildren.length > 0 ||
        firstHeaderChildren.length > 0
          ? {
              headers: {
                ...(defaultHeaderChildren.length > 0 ? { default: defaultHeaderChildren } : {}),
                ...(evenHeaderChildren.length > 0 ? { even: evenHeaderChildren } : {}),
                ...(firstHeaderChildren.length > 0 ? { first: firstHeaderChildren } : {}),
              },
            }
          : {}),
        ...(defaultFooterChildren.length > 0 ||
        evenFooterChildren.length > 0 ||
        firstFooterChildren.length > 0
          ? {
              footers: {
                ...(defaultFooterChildren.length > 0 ? { default: defaultFooterChildren } : {}),
                ...(evenFooterChildren.length > 0 ? { even: evenFooterChildren } : {}),
                ...(firstFooterChildren.length > 0 ? { first: firstFooterChildren } : {}),
              },
            }
          : {}),
      },
    ],
    ...(stylesheet.styles.length > 0
      ? {
          styles: {
            paragraphStyles: stylesheet.styles
              .filter((style) => style.type === "paragraph")
              .map((style) => ({ id: style.id, name: style.name, basedOn: style.basedOn })),
            characterStyles: stylesheet.styles
              .filter((style) => style.type === "character")
              .map((style) => ({ id: style.id, name: style.name, basedOn: style.basedOn })),
          },
        }
      : {}),
    ...(lists && (lists.abstractNumberings.length > 0 || lists.overrides.length > 0)
      ? {
          numbering: {
            abstractNumberings: lists.abstractNumberings.map((numbering) => ({
              reference: numbering.reference,
              levels: numbering.levels,
            })),
          },
        }
      : {}),
    ...(fib.ccpFootnotes > 0
      ? {
          footnotes: parseOptional(
            () => parseNoteBoundaries(table, fib.footnoteTable, fib.ccpFootnotes),
            [],
          ).map((boundary, index) => ({
            id: index + 1,
            children: projectStory(
              characters,
              characterRanges,
              footnoteStart + boundary.start,
              footnoteStart + boundary.end,
            ),
          })),
        }
      : {}),
    ...(fib.ccpComments > 0
      ? {
          comments: [
            {
              id: 1,
              children: projectCommentStory(
                characters,
                characterRanges,
                commentStart,
                endnoteStart,
              ),
            },
          ],
        }
      : {}),
    ...(fib.ccpEndnotes > 0
      ? {
          endnotes: parseOptional(
            () => parseNoteBoundaries(table, fib.endnoteTable, fib.ccpEndnotes),
            [],
          ).map((boundary, index) => ({
            id: index + 1,
            children: projectStory(
              characters,
              characterRanges,
              endnoteStart + boundary.start,
              endnoteStart + boundary.end,
            ),
          })),
        }
      : {}),
  };
  Object.assign(result, metadata);
  if (fields.length > 0) result.fields = fields;
  if (revisions.length > 0) result.revisions = revisions;
  if (pictures.length > 0) result.pictures = pictures;
  return result;
}

function applyBookmarks(
  children: SectionChild[],
  bookmarks: readonly { id: number; name: string; start: number; end: number }[],
  characters: readonly TextCharacter[],
): SectionChild[] {
  const boundaries = new Map<number, Array<{ marker: SectionChild; before: boolean }>>();
  for (const bookmark of bookmarks) {
    const startIndex = paragraphIndexForCp(characters, bookmark.start);
    boundaries.set(startIndex, [
      ...(boundaries.get(startIndex) ?? []),
      { marker: { bookmarkStart: { id: bookmark.id, name: bookmark.name } }, before: true },
    ]);
    const endIndex = paragraphIndexForCp(characters, bookmark.end);
    boundaries.set(endIndex, [
      ...(boundaries.get(endIndex) ?? []),
      { marker: { bookmarkEnd: { id: bookmark.id } }, before: false },
    ]);
  }
  if (boundaries.size === 0) return children;
  return children.flatMap((child, index) => {
    const markers = boundaries.get(index) ?? [];
    return [
      ...markers.filter((item) => item.before).map((item) => item.marker),
      child,
      ...markers.filter((item) => !item.before).map((item) => item.marker),
    ];
  });
}

function parseCommentRanges(
  table: Uint8Array,
  range: NumberPair,
): Array<{ start: number; end: number; data: Uint8Array }> {
  const message = "Invalid DOC annotation references: they are outside the table stream";
  requireRange(table, range.offset, range.length, message);
  if (range.length < 8 || (range.length - 4) % 4 !== 0) {
    throw new DocParseError("Invalid DOC annotation references: malformed CP count");
  }
  const count = (range.length - 4) / 4;
  const values: Array<{ start: number; end: number; data: Uint8Array }> = [];
  for (let index = 0; index < count; index += 1) {
    const start = readUint32(table, range.offset + index * 4, message);
    const end = readUint32(table, range.offset + (index + 1) * 4, message);
    const dataOffset = range.offset + (count + 1) * 4 + index * 4;
    values.push({
      start,
      end,
      data: table.subarray(dataOffset, dataOffset + 4),
    });
  }
  return values;
}
