import { CompoundFileReader } from "@office-open/core";
import type {
  DocumentOptions,
  ParagraphOptions,
  RunOptions,
  SectionChild,
} from "@office-open/docx";

import { DocParseError } from "./errors";

interface Fib {
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
}

interface NumberPair {
  offset: number;
  length: number;
}

interface Piece {
  cpStart: number;
  cpEnd: number;
  fc: number;
  compressed: boolean;
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
}

interface TextCharacter {
  value: string;
  cp: number;
  paragraphEnd?: boolean;
}

const FLAG_ENCRYPTED = 0x0100;
const FLAG_TABLE_ONE = 0x0200;
const FIB_SIGNATURE = 0xa5ec;
const FIB_BASE_LENGTH = 32;
const UNICODE_PIECE = 0x40000000;
const PIECE_OFFSET_MASK = 0x3fffffff;
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
    throw new DocParseError(message);
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
    throw new DocParseError("Invalid Word document: FIB signature is not 0xA5EC");
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
      throw new DocParseError(`Invalid Word document: ${name} is outside the FIB`);
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
    characterBinTable: readPair(12, "fcPlcfbteChpx"),
    paragraphBinTable: readPair(13, "fcPlcfbtePapx"),
    footnoteTable: readPair(3, "fcPlcffndTxt"),
    headerTable: readPair(11, "fcPlcfHdd"),
    endnoteTable: readPair(47, "fcPlcfendTxt"),
    clx: readPair(33, "fcClx"),
    textboxTable: readPair(56, "fcPlcftxbxTxt"),
    headerTextboxTable: readPair(58, "fcPlcfHdrTxbxTxt"),
  };
}

function parsePieceTable(table: Uint8Array, clx: NumberPair): Piece[] {
  const rangeMessage = "Invalid DOC CLX: piece table is outside the table stream";
  requireRange(table, clx.offset, clx.length, rangeMessage);
  if (clx.length === 0)
    throw new DocParseError("Invalid Word document: CLX piece table is missing");

  let position = clx.offset;
  const end = clx.offset + clx.length;
  while (position < end) {
    const kind = table[position]!;
    if (kind === 1) {
      requireRange(table, position, 3, "Invalid DOC CLX: truncated property record");
      position += 3 + table[position + 1]! + (table[position + 2]! << 8);
      continue;
    }
    if (kind === 2) {
      requireRange(table, position, 5, "Invalid DOC CLX: truncated piece table header");
      const length = readUint32(table, position + 1, "Invalid DOC CLX: invalid piece table length");
      const tableStart = position + 5;
      requireRange(table, tableStart, length, rangeMessage);
      if (tableStart + length !== end) {
        throw new DocParseError("Invalid DOC CLX: piece table does not terminate CLX");
      }
      if (length < 12 || (length - 4) % 12 !== 0) {
        throw new DocParseError("Invalid DOC CLX: malformed piece table size");
      }
      return parsePieces(table, tableStart, length);
    }
    throw new DocParseError(`Invalid DOC CLX: unknown record type ${kind}`);
  }
  throw new DocParseError("Invalid DOC CLX: piece table marker is missing");
}

function parsePieces(table: Uint8Array, start: number, length: number): Piece[] {
  const pieceCount = (length - 4) / 12;
  const pieces: Piece[] = [];
  for (let index = 0; index < pieceCount; index++) {
    const cpOffset = start + index * 4;
    const cpStart = readUint32(
      table,
      cpOffset,
      "Invalid DOC piece table: CP is outside the piece table",
    );
    const cpEnd = readUint32(
      table,
      cpOffset + 4,
      "Invalid DOC piece table: CP is outside the piece table",
    );
    if (index === 0 && cpStart !== 0) {
      throw new DocParseError("Invalid DOC piece table: first CP is not zero");
    }
    if (cpEnd <= cpStart || (index > 0 && cpStart !== pieces[index - 1]!.cpEnd)) {
      throw new DocParseError("Invalid DOC piece table: CP boundaries are not contiguous");
    }

    const pcdOffset = start + (pieceCount + 1) * 4 + index * 8;
    const fcValue = readUint32(
      table,
      pcdOffset + 2,
      "Invalid DOC piece table: PCD is outside the piece table",
    );
    const { fc, compressed } = transformPieceFc(fcValue);
    if (!Number.isSafeInteger(fc) || fc < 0) {
      throw new DocParseError("Invalid DOC piece table: invalid compressed FC");
    }
    pieces.push({ cpStart, cpEnd, fc, compressed });
  }
  return pieces;
}

function transformPieceFc(value: number): { fc: number; compressed: boolean } {
  const compressed = (value & UNICODE_PIECE) !== 0;
  return {
    compressed,
    fc: compressed ? (value & PIECE_OFFSET_MASK) / 2 : value & PIECE_OFFSET_MASK,
  };
}

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

function parseHeaderStories(
  table: Uint8Array,
  range: NumberPair,
  storyLength: number,
): TextboxBoundary[] {
  if (storyLength === 0) return [];
  if (range.length === 0) {
    throw new DocParseError("Invalid Word document: header boundary table is missing");
  }
  const message = "Invalid DOC header table: it is outside the table stream";
  requireRange(table, range.offset, range.length, message);
  if (range.length < 12 || range.length % 4 !== 0) {
    throw new DocParseError("Invalid DOC header table: malformed CP count");
  }
  const cps = Array.from({ length: range.length / 4 }, (_, index) =>
    readUint32(table, range.offset + index * 4, message),
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
  if (cps[0] !== 0 || cps[count - 1]! > storyLength) {
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
  if (count === 0 || count > (FKP_LENGTH - 5) / 14) {
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
      ranges.push({ cpStart: fcStart, cpEnd: fcEnd });
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
    skipProperties(word.subarray(grpprlStart, grpprlStart + propertiesLength));
    ranges.push({ cpStart: fcStart, cpEnd: fcEnd });
  }
  return ranges;
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

function skipProperties(grpprl: Uint8Array): void {
  iterateProperties(grpprl);
}

function parseCharacterProperties(grpprl: Uint8Array): RunOptions {
  const properties: RunOptions = {};
  iterateProperties(grpprl, (opcode, operandOffset, length) => {
    if (opcode === 0x0835 && length === 1)
      properties.bold = grpprl[operandOffset]! === 1 || grpprl[operandOffset]! === 128;
    if (opcode === 0x0836 && length === 1)
      properties.italic = grpprl[operandOffset]! === 1 || grpprl[operandOffset]! === 128;
    if (opcode === 0x0837 && length === 1)
      properties.strike = grpprl[operandOffset]! === 1 || grpprl[operandOffset]! === 128;
    if (opcode === 0x4a43 && length === 2) {
      properties.size = readUint16(grpprl, operandOffset, "Invalid DOC character properties") / 2;
    }
    if (opcode === 0x6a03 && length === 2) {
      const colorIndex = grpprl[operandOffset]!;
      if (colorIndex < COLOR_INDEXES.length) properties.color = COLOR_INDEXES[colorIndex]!;
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
): SectionChild[] {
  const children: SectionChild[] = [];
  let paragraph: TextCharacter[] = [];

  const flushParagraph = (): void => {
    if (paragraph.length === 0) {
      children.push({ paragraph: "" });
      return;
    }
    const boundaries = new Set<number>([0, paragraph.length]);
    for (let index = 0; index < paragraph.length; index++) {
      const cp = paragraph[index]!.cp;
      for (const range of ranges) {
        if (cp === range.cpStart) boundaries.add(index);
      }
    }
    const sortedBoundaries = [...boundaries].sort((left, right) => left - right);
    const runs: RunOptions[] = [];
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
      const breakCount = text.split("\n").length - 1;
      const visibleText = text.replaceAll("\n", "");
      if (breakCount > 0) runs.push({ break: breakCount });
      if (visibleText.length > 0) runs.push({ text: visibleText, ...range?.properties });
    }
    const plainText = paragraph.map((character) => character.value).join("");
    const hasProperties = runs.some((run) =>
      Object.keys(run).some((key) => key !== "text" && key !== "break"),
    );
    children.push({ paragraph: hasProperties || runs.length > 1 ? { children: runs } : plainText });
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
): SectionChild[] {
  return projectChildren(
    characters.filter((character) => character.cp >= start && character.cp < end),
    ranges,
  );
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

export function parseInternal(data: Uint8Array): DocumentOptions {
  let reader: CompoundFileReader;
  try {
    reader = new CompoundFileReader(data);
  } catch (error) {
    throw new DocParseError(
      `Input is not a supported Compound File Binary document: ${(error as Error).message}`,
    );
  }
  if (!reader.entry("WordDocument")) {
    throw new DocParseError("Invalid Word document: WordDocument stream is missing");
  }
  const word = reader.read("WordDocument");
  const fib = parseFib(word);
  const tablePath =
    (readUint16(word, 10, "Invalid Word document: FIB is truncated") & FLAG_TABLE_ONE) !== 0
      ? "1Table"
      : "0Table";
  if (!reader.entry(tablePath)) {
    throw new DocParseError(`Invalid Word document: ${tablePath} stream is missing`);
  }
  const table = reader.read(tablePath);
  if (fib.totalCharacters === 0) return { sections: [{ children: [{ paragraph: "" }] }] };

  const pieces = parsePieceTable(table, fib.clx);
  if (fib.totalCharacters > pieces[pieces.length - 1]!.cpEnd) {
    throw new DocParseError("Invalid Word document: main text extends beyond the piece table");
  }
  const characters = projectText(pieces, word, fib.totalCharacters);
  const characterEntries = parseBinTable(table, fib.characterBinTable, "character");
  const characterRanges = characterEntries.flatMap((entry) =>
    parseCharacterFkp(word, entry.page, pieces, fib.totalCharacters),
  );
  for (const entry of parseBinTable(table, fib.paragraphBinTable, "paragraph")) {
    parseParagraphFkp(word, entry.page);
  }

  const ranges = [
    { start: 0, end: fib.ccpText },
    { start: fib.ccpText, end: fib.ccpText + fib.ccpFootnotes },
    {
      start: fib.ccpText + fib.ccpFootnotes,
      end: fib.ccpText + fib.ccpFootnotes + fib.ccpHeaders,
    },
  ];
  const footnoteStart = ranges[1]!.start;
  const headerStart = ranges[2]!.start;
  const commentStart = headerStart + fib.ccpHeaders;
  const endnoteStart = commentStart + fib.ccpComments;
  const textboxStart = endnoteStart + fib.ccpEndnotes;
  const headerTextboxStart = textboxStart + fib.ccpTextboxes;

  const sectionChildren = projectStory(characters, characterRanges, 0, fib.ccpText);
  const headerStories = parseHeaderStories(table, fib.headerTable, fib.ccpHeaders);
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
  for (const boundary of fib.ccpTextboxes > 0
    ? parseTextboxBoundaries(table, fib.textboxTable, fib.ccpTextboxes)
    : []) {
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
  for (const boundary of fib.ccpHeaderTextboxes > 0
    ? parseTextboxBoundaries(table, fib.headerTextboxTable, fib.ccpHeaderTextboxes)
    : []) {
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

  return {
    sections: [
      {
        children: sectionChildren,
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
    ...(fib.ccpFootnotes > 0
      ? {
          footnotes: parseNoteBoundaries(table, fib.footnoteTable, fib.ccpFootnotes).map(
            (boundary, index) => ({
              id: index + 1,
              children: projectStory(
                characters,
                characterRanges,
                footnoteStart + boundary.start,
                footnoteStart + boundary.end,
              ),
            }),
          ),
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
          endnotes: parseNoteBoundaries(table, fib.endnoteTable, fib.ccpEndnotes).map(
            (boundary, index) => ({
              id: index + 1,
              children: projectStory(
                characters,
                characterRanges,
                endnoteStart + boundary.start,
                endnoteStart + boundary.end,
              ),
            }),
          ),
        }
      : {}),
  };
}

export { transformPieceFc };
