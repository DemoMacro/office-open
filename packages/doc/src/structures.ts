import { DocParseError } from "./errors";
import type { LegacyBookmark, LegacyField, LegacyPictureData } from "./types";

interface NumberPair {
  offset: number;
  length: number;
}

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

interface PlcRange {
  start: number;
  end: number;
  data: Uint8Array;
}

function parsePlcf(
  table: Uint8Array,
  range: NumberPair,
  recordLength: number,
  message: string,
): PlcRange[] {
  requireRange(table, range.offset, range.length, message);
  if (range.length < 8 || (range.length - 4) % (4 + recordLength) !== 0) {
    throw new DocParseError(`${message}: malformed CP count`);
  }
  const count = (range.length - 4) / (4 + recordLength);
  const values: PlcRange[] = [];
  for (let index = 0; index < count; index += 1) {
    const start = readUint32(table, range.offset + index * 4, message);
    const end = readUint32(table, range.offset + (index + 1) * 4, message);
    const dataOffset = range.offset + (count + 1) * 4 + index * recordLength;
    values.push({ start, end, data: table.subarray(dataOffset, dataOffset + recordLength) });
  }
  return values;
}

function decodeWindows1252(bytes: Uint8Array): string {
  let result = "";
  for (const byte of bytes) result += WINDOWS_1252_HIGH.get(byte) ?? String.fromCharCode(byte);
  return result;
}

function readUnicodeString(bytes: Uint8Array, offset: number): string {
  if (offset + 2 > bytes.byteLength) return "";
  const length = readUint16(bytes, offset, "Invalid DOC stylesheet: truncated string");
  const start = offset + 2;
  requireRange(bytes, start, length * 2, "Invalid DOC stylesheet: truncated string");
  let result = "";
  for (let index = 0; index < length; index += 1) {
    result += String.fromCharCode(readUint16(bytes, start + index * 2, "Invalid DOC stylesheet"));
  }
  return result;
}

/** Parse Stshf (style sheet): style names, style IDs and based-on relationships. */
export function parseStylesheet(table: Uint8Array, range: NumberPair): StylesheetModel {
  const message = "Invalid DOC stylesheet: it is outside the table stream";
  requireRange(table, range.offset, range.length, message);
  if (range.length < 4) throw new DocParseError("Invalid DOC stylesheet: malformed STSH");
  const count = readUint16(table, range.offset, message);
  let offset = range.offset + 4;
  const styles: StyleModel[] = [];
  for (let index = 0; index < count && offset + 12 <= range.offset + range.length; index += 1) {
    const type = table[offset]!;
    const istd = readUint16(table, offset + 2, message);
    const base = readUint16(table, offset + 4, message);
    const nameOffset = readUint16(table, offset + 6, message);
    const name = readUnicodeString(table, offset + nameOffset);
    styles.push({
      type: type === 2 ? "character" : type === 3 ? "table" : "paragraph",
      id: `style-${istd}`,
      styleId: istd,
      name: name || `style ${istd}`,
      basedOn: base === 4095 || base === istd ? undefined : `style-${base}`,
    });
    const nameLength = readUint16(table, offset + nameOffset, message);
    offset += nameOffset + 2 + nameLength * 2;
    if (offset > range.offset + range.length) {
      throw new DocParseError("Invalid DOC stylesheet: style description is truncated");
    }
  }
  return { styles };
}

export interface StyleModel {
  type: "paragraph" | "character" | "table";
  id: string;
  styleId: number;
  name: string;
  basedOn?: string;
}

export interface StylesheetModel {
  styles: StyleModel[];
}

const ALIGNMENTS = ["left", "center", "right", "both"] as const;

/** Parse SEPX records from PlcfSed into public section geometry. */
export function parseSectionProperties(
  table: Uint8Array,
  plcfsepx: NumberPair,
  word: Uint8Array,
): SectionPropertiesModel {
  const message = "Invalid DOC section table: it is outside the table stream";
  const records = parsePlcf(table, plcfsepx, 12, message);
  const properties: SectionModel[] = [];
  for (const record of records) {
    try {
      parseOneSection(record, word, properties);
    } catch {
      // Tolerate individual corrupt SEPX records — sections are auxiliary
      // geometry and the remaining sections should still be read.
    }
  }
  return { sections: properties };
}

function parseOneSection(
  record: { data: Uint8Array },
  word: Uint8Array,
  properties: SectionModel[],
): void {
  const sepx = readUint32(record.data, 4, "Invalid DOC section properties: invalid SEPX pointer");
  const cb = readUint16(word, sepx, "Invalid DOC section properties: truncated SEPX");
  const grpprlStart = sepx + 2;
  requireRange(word, grpprlStart, cb, "Invalid DOC section properties: truncated SEPX");
  const section: SectionModel = {};
  for (let index = 0; index + 2 <= cb;) {
    const opcode = readUint16(word, grpprlStart + index, "Invalid DOC section properties");
    const form = (opcode >> 13) & 7;
    const operation = opcode & 0x00ff;
    const length =
      operation === 0x1d
        ? 1
        : operation >= 0x1f && operation <= 0x24
          ? 2
          : form === 0 || form === 1
            ? 1
            : form === 3
              ? 4
              : 2;
    if (index + 2 + length > cb) {
      throw new DocParseError("Invalid DOC section properties: operand is truncated");
    }
    const operand = grpprlStart + index + 2;
    if (operation === 0x1d && length >= 1) {
      section.orientation = word[operand] === 1 ? "landscape" : "portrait";
    } else if (operation === 0x1f && length >= 2) {
      section.width = word[operand]! | (word[operand + 1]! << 8);
    } else if (operation === 0x20 && length >= 2) {
      section.height = word[operand]! | (word[operand + 1]! << 8);
    } else if (operation === 0x21 && length >= 2) {
      section.marginLeft = word[operand]! | (word[operand + 1]! << 8);
    } else if (operation === 0x22 && length >= 2) {
      section.marginRight = word[operand]! | (word[operand + 1]! << 8);
    } else if (operation === 0x23 && length >= 2) {
      section.marginTop = word[operand]! | (word[operand + 1]! << 8);
    } else if (operation === 0x24 && length >= 2) {
      section.marginBottom = word[operand]! | (word[operand + 1]! << 8);
    }
    index += 2 + length;
  }
  properties.push(section);
}

export interface SectionModel {
  width?: number;
  height?: number;
  orientation?: "portrait" | "landscape";
  marginTop?: number;
  marginRight?: number;
  marginBottom?: number;
  marginLeft?: number;
  alignment?: (typeof ALIGNMENTS)[number];
}

export interface SectionPropertiesModel {
  sections: SectionModel[];
}

/** Parse PlfLst/PlfLfo into abstract numbering and list overrides. */
export function parseLists(table: Uint8Array, list: NumberPair, overrides: NumberPair): ListModel {
  const message = "Invalid DOC list table: it is outside the table stream";
  const abstractNumberings: AbstractNumberingModel[] = [];
  if (list.length >= 2) {
    requireRange(table, list.offset, list.length, message);
    const count = readUint16(table, list.offset, message);
    let offset = list.offset + 2;
    for (let index = 0; index < count; index += 1) {
      requireRange(table, offset, 28, message);
      const reference = readUint32(table, offset, message);
      const flags = table[offset + 26]!;
      const levelCount = (flags & 0x01) === 0 ? 9 : 1;
      abstractNumberings.push({
        reference: `list-${reference}`,
        simple: (flags & 0x01) !== 0,
        levels: Array.from({ length: levelCount }, (_, level) => ({
          level,
          format: "decimal",
          text: `%${level + 1}.`,
          start: 1,
        })),
      });
      offset += 28;
    }
  }
  const listOverrides: ListOverrideModel[] = [];
  if (overrides.length >= 4) {
    requireRange(table, overrides.offset, overrides.length, message);
    const count = readUint32(table, overrides.offset, message);
    let offset = overrides.offset + 4;
    for (let index = 0; index < count; index += 1) {
      requireRange(table, offset, 6, message);
      listOverrides.push({
        listIndex: index,
        reference: readUint32(table, offset, message),
        overrideCount: table[offset + 4]!,
      });
      offset += 6;
    }
  }
  return { abstractNumberings, overrides: listOverrides };
}

export interface LevelModel {
  level: number;
  format: "decimal";
  text: string;
  start: number;
}

export interface AbstractNumberingModel {
  reference: string;
  simple: boolean;
  levels: LevelModel[];
}

export interface ListOverrideModel {
  listIndex: number;
  reference: number;
  overrideCount: number;
}

export interface ListModel {
  abstractNumberings: AbstractNumberingModel[];
  overrides: ListOverrideModel[];
}

/** Parse bookmark name table plus start/end PLCFs. */
export function parseBookmarks(
  table: Uint8Array,
  names: NumberPair,
  starts: NumberPair,
  ends: NumberPair,
): LegacyBookmark[] {
  const nameMessage = "Invalid DOC bookmark names: they are outside the table stream";
  requireRange(table, names.offset, names.length, nameMessage);
  if (names.length < 6) throw new DocParseError("Invalid DOC bookmark names: malformed STTB");
  const count = readUint16(table, names.offset, nameMessage);
  let cursor = names.offset + 4;
  const bookmarkNames: string[] = [];
  for (let index = 0; index < count; index += 1) {
    const length = readUint16(table, cursor, nameMessage);
    cursor += 2;
    requireRange(table, cursor, length, nameMessage);
    bookmarkNames.push(decodeWindows1252(table.subarray(cursor, cursor + length)));
    cursor += length;
  }
  const startsByCp = parsePlcf(table, starts, 4, "Invalid DOC bookmark start table");
  requireRange(table, ends.offset, ends.length, "Invalid DOC bookmark end table");
  if (ends.length === 0 || ends.length % 4 !== 0) {
    throw new DocParseError("Invalid DOC bookmark end table: malformed CP count");
  }
  const endCps = Array.from({ length: ends.length / 4 }, (_, index) =>
    readUint32(table, ends.offset + index * 4, "Invalid DOC bookmark end table"),
  );
  return startsByCp.map((start, index) => ({
    id: index + 1,
    name: bookmarkNames[index] ?? `bookmark-${index + 1}`,
    start: start.start,
    end: endCps[index + 1] ?? start.end,
  }));
}

/** Parse PlcffldMom and recover field instruction bytes between CHRs 19/20. */
export function parseFields(
  characters: readonly string[],
  table: Uint8Array,
  range: NumberPair,
): LegacyField[] {
  const message = "Invalid DOC field table: it is outside the table stream";
  const records = parsePlcf(table, range, 2, message);
  const fields: LegacyField[] = [];
  for (const record of records) {
    const kind = record.data[0];
    if (kind === 0x13) {
      const cp = record.start;
      let end = Math.min(cp + 255, characters.length);
      for (let index = cp + 1; index < end; index += 1) {
        if (characters[index] === "\x14") {
          end = index;
          break;
        }
      }
      const instruction = characters
        .slice(cp + 1, end)
        .join("")
        .replace(/[\x13\x14\x15]/g, "");
      fields.push({ start: cp, end, instruction });
    }
  }
  return fields;
}

/** Parse Escher OfficeArtContent, including BStore containers and BLIP records. */
export function parsePictures(drawing: Uint8Array, dataStream: Uint8Array): LegacyPictureData[] {
  const pictures: LegacyPictureData[] = [];
  const walk = (offset: number, end: number): void => {
    while (offset + 8 <= end) {
      const version = readUint16(drawing, offset, "Invalid DOC Escher drawing") & 0x000f;
      const type = readUint16(drawing, offset + 2, "Invalid DOC Escher drawing");
      const length = readUint32(drawing, offset + 4, "Invalid DOC Escher drawing");
      const body = offset + 8;
      if (length > end - body) return;
      if (version === 0x0f) {
        walk(body, body + length);
      } else if (type === 0xf007) {
        const blipReference = readUint32(drawing, body + 20, "Invalid DOC Escher picture");
        const blipType = drawing[body + 10];
        if (blipReference !== 0xffffffff && blipType !== undefined) {
          pictures.push(parseBlip(dataStream, blipReference));
        }
      } else if (type === 0xf01a || type === 0xf01b || type === 0xf01c || type === 0xf00a) {
        pictures.push(parseEmbeddedBlip(drawing.subarray(body, body + length)));
      }
      offset = body + length;
    }
  };
  walk(0, drawing.byteLength);
  return pictures.filter((picture) => picture.data.byteLength > 0);
}

const BLIP_TYPES = new Map<number, LegacyPictureData["type"]>([
  [0x2160, "bmp"],
  [0x6c00, "jpeg"],
  [0x6c01, "png"],
  [0x7c80, "gif"],
  [0x6c02, "tiff"],
]);

function imageTypeFromBytes(data: Uint8Array): LegacyPictureData["type"] | undefined {
  if (data[0] === 0x89 && data[1] === 0x50 && data[2] === 0x4e && data[3] === 0x47) return "png";
  if (data[0] === 0xff && data[1] === 0xd8) return "jpeg";
  if (data[0] === 0x47 && data[1] === 0x49 && data[2] === 0x46) return "gif";
  if (data[0] === 0x42 && data[1] === 0x4d) return "bmp";
  if ((data[0] === 0x49 && data[1] === 0x20) || (data[0] === 0x4d && data[1] === 0x4d))
    return "tiff";
  return undefined;
}

function parseEmbeddedBlip(bytes: Uint8Array): LegacyPictureData {
  const record = readUint16(bytes, 2, "Invalid DOC Escher BLIP");
  let start = 36;
  if (bytes[0] === 0x89 && bytes[1] === 0x50) start = 0;
  else if (bytes[0] === 0xff && bytes[1] === 0xd8) start = 0;
  const type = imageTypeFromBytes(bytes.subarray(start)) ?? BLIP_TYPES.get(record) ?? "unknown";
  const data = bytes.subarray(start);
  const dimensions = readImageDimensions(data, type);
  return { type, data, ...dimensions };
}

function parseBlip(dataStream: Uint8Array, reference: number): LegacyPictureData {
  let offset = 0;
  while (offset + 61 <= dataStream.byteLength) {
    if (readUint32(dataStream, offset + 24, "Invalid DOC Data stream") === reference) {
      return parseEmbeddedBlip(dataStream.subarray(offset + 36));
    }
    offset += 61;
  }
  return { type: "unknown", data: new Uint8Array(), width: 0, height: 0 };
}

function readImageDimensions(
  data: Uint8Array,
  type: LegacyPictureData["type"],
): { width: number; height: number } {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  try {
    if (type === "png" && data.byteLength > 23) {
      return { width: view.getUint32(16), height: view.getUint32(20) };
    }
    if (type === "jpeg") {
      for (let offset = 2; offset + 8 < data.byteLength;) {
        if (data[offset] !== 0xff) break;
        const marker = data[offset + 1]!;
        const length = view.getUint16(offset + 2);
        if (
          marker >= 0xc0 &&
          marker <= 0xcf &&
          marker !== 0xc4 &&
          marker !== 0xc8 &&
          marker !== 0xcc
        ) {
          return { height: view.getUint16(offset + 5), width: view.getUint16(offset + 7) };
        }
        offset += 2 + length;
      }
    }
    if (type === "gif" && data.byteLength > 9) {
      return { width: view.getUint16(6, true), height: view.getUint16(8, true) };
    }
    if (type === "bmp" && data.byteLength > 25) {
      return {
        width: Math.abs(view.getInt32(18, true)),
        height: Math.abs(view.getInt32(22, true)),
      };
    }
  } catch {
    return { width: 0, height: 0 };
  }
  return { width: 0, height: 0 };
}
