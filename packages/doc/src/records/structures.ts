import {
  readUint16 as readStreamUint16,
  readUint32 as readStreamUint32,
  requireByteRange,
} from "../binary";
import { DocParseError, type DocParseContext } from "../errors";
import type {
  LegacyBookmark,
  LegacyField,
  LegacyPictureByShapeId,
  LegacyPictureData,
} from "./models";

interface NumberPair {
  offset: number;
  length: number;
}
const TABLE_LOCATION = { part: "table", path: "table" } as const;

function requireRange(
  bytes: Uint8Array,
  offset: number,
  length: number,
  message: string,
  context: Partial<Omit<DocParseContext, "format">> = {},
): void {
  requireByteRange(bytes, offset, length, message, context, TABLE_LOCATION);
}

function readUint16(bytes: Uint8Array, offset: number, message: string): number {
  return readStreamUint16(bytes, offset, message, {}, TABLE_LOCATION);
}

function readUint32(bytes: Uint8Array, offset: number, message: string): number {
  return readStreamUint32(bytes, offset, message, {}, TABLE_LOCATION);
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
    throw new DocParseError(`${message}: malformed CP count`, {
      part: "table",
      path: "PLC",
      offset: range.offset,
      length: range.length,
      byteRange: [range.offset, range.offset + range.length],
      reason: "invalid-record-length",
    });
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
  if (range.length < 6) throw new DocParseError("Invalid DOC stylesheet: malformed STSH");
  const headerLength = readUint16(table, range.offset, message);
  if (headerLength < 6) throw new DocParseError("Invalid DOC stylesheet: malformed STSHI");
  const baseLength = readUint16(table, range.offset + 4, message);
  if (baseLength < 8) throw new DocParseError("Invalid DOC stylesheet: malformed STD base");
  const count = readUint16(table, range.offset + 2, message);
  let offset = range.offset + 2 + headerLength;
  const styles: StyleModel[] = [];
  const end = range.offset + range.length;
  for (let index = 0; index < count; index += 1) {
    if (offset + 2 > end) {
      throw new DocParseError("Invalid DOC stylesheet: style record is truncated");
    }
    const recordLength = readUint16(table, offset, message);
    if (recordLength === 0) continue;
    if (recordLength % 2 !== 0 || offset + 2 + recordLength > end) {
      throw new DocParseError("Invalid DOC stylesheet: style description is truncated");
    }
    if (recordLength < 8) {
      throw new DocParseError("Invalid DOC stylesheet: style description is truncated");
    }

    const record = table.subarray(offset + 2, offset + 2 + recordLength);
    const info2 = readUint16(record, 2, message);
    const styleKind = info2 & 0x0f;
    const basedOn = (info2 >> 4) & 0x0fff;
    const nameOffset = Math.min(baseLength, recordLength);
    if (nameOffset + 2 > recordLength) {
      throw new DocParseError("Invalid DOC stylesheet: style name is truncated");
    }
    const nameLength = readUint16(record, nameOffset, message);
    if (nameOffset + 2 + nameLength * 2 > recordLength) {
      throw new DocParseError("Invalid DOC stylesheet: style name is truncated");
    }
    const rawName = readUnicodeString(record, nameOffset);
    const name = (nameLength > 0 && rawName.endsWith("\0") ? rawName.slice(0, -1) : rawName).split(
      ",",
      1,
    )[0]!;
    let type: StyleModel["type"] = "paragraph";
    if (styleKind === 2) type = "character";
    else if (styleKind === 3) type = "table";
    styles.push({
      type,
      id: `style-${index}`,
      styleId: index,
      name: name || `style ${index}`,
      basedOn: basedOn === 4095 || basedOn === index ? undefined : `style-${basedOn}`,
    });
    offset += 2 + recordLength;
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
    parseOneSection(record, word, properties);
  }
  return { sections: properties };
}

function parseOneSection(
  record: { start: number; end: number; data: Uint8Array },
  word: Uint8Array,
  properties: SectionModel[],
): void {
  const sepx = readUint32(record.data, 2, "Invalid DOC section properties: invalid SEPX pointer");
  const cb = readUint16(word, sepx, "Invalid DOC section properties: truncated SEPX");
  const grpprlStart = sepx + 2;
  requireRange(word, grpprlStart, cb, "Invalid DOC section properties: truncated SEPX");
  const section: SectionModel = {};
  section.mainStart = record.start;
  section.mainEnd = record.end;
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
  /** Main-document CP range start projected into this section. */
  mainStart?: number;
  /** Main-document CP range end projected into this section. */
  mainEnd?: number;
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
      if (offset + 28 > list.offset + list.length) {
        throw new DocParseError(`${message}: truncated abstract numbering`, {
          part: "table",
          path: "PlfLst",
          offset,
          length: 28,
          reason: "out-of-range",
        });
      }
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
      if (offset + 6 > overrides.offset + overrides.length) {
        throw new DocParseError(`${message}: truncated list override`, {
          part: "table",
          path: "Plcflfo",
          offset,
          length: 6,
          reason: "out-of-range",
        });
      }
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
  if (readUint16(table, names.offset, nameMessage) !== 0xffff) {
    throw new DocParseError("Invalid DOC bookmark names: non-extended STTB");
  }
  const count = readUint16(table, names.offset + 2, nameMessage);
  if (readUint16(table, names.offset + 4, nameMessage) !== 0) {
    throw new DocParseError("Invalid DOC bookmark names: unexpected extra data");
  }
  let cursor = names.offset + 6;
  const bookmarkNames: string[] = [];
  for (let index = 0; index < count; index += 1) {
    const length = readUint16(table, cursor, nameMessage);
    cursor += 2;
    requireRange(table, cursor, length * 2, nameMessage);
    bookmarkNames.push(readUnicodeString(table, cursor - 2));
    cursor += length * 2;
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
  if (range.length === 4) return [];
  const records = parsePlcf(table, range, 2, message);
  const fields: LegacyField[] = [];
  for (const record of records) {
    if ((record.data[0]! & 0x7f) === 0x13) {
      const cp = record.start;
      const limit = Math.min(cp + 255, characters.length);
      let separator = -1;
      for (let scan = cp + 1; scan < limit; scan += 1) {
        if (characters[scan] === "\x14") {
          separator = scan;
          break;
        }
      }
      if (separator < 0) {
        for (let scan = cp + 1; scan < limit; scan += 1) {
          if (characters[scan] === "\x15") {
            fields.push({
              start: cp,
              end: scan,
              instruction: characters.slice(cp + 1, scan).join(""),
            });
            break;
          }
        }
        if (fields.at(-1)?.start === cp) continue;
        throw new DocParseError("Invalid DOC field table: field has no separator", {
          part: "stream",
          path: "WordDocument",
          recordName: "ComplexField",
          offset: cp,
          reason: "missing-required-record",
        });
      }
      fields.push({
        start: cp,
        end: separator,
        instruction: characters.slice(cp + 1, separator).join(""),
      });
    }
  }
  return fields;
}

/** Parse Escher OfficeArtContent and resolve picture-bearing shapes by shape id. */
export function parsePictures(drawing: Uint8Array, dataStream: Uint8Array): LegacyPictureByShapeId {
  const pictures: LegacyPictureData[] = [];
  const picturesByShapeId: LegacyPictureByShapeId = new Map();

  const isEscherType = (type: number): boolean =>
    (type >= 0xf000 && type <= 0xf012) ||
    type === 0xf11e ||
    type === 0xf121 ||
    type === 0xf122 ||
    (type >= 0xf01a && type <= 0xf01f);

  const alignRecord = (offset: number, end: number, path: string): number => {
    const canRead = (candidate: number): boolean => candidate + 8 <= end;
    const read = (candidate: number) => ({
      version: readUint16(drawing, candidate, "Invalid DOC Escher drawing") & 0x000f,
      type: readUint16(drawing, candidate + 2, "Invalid DOC Escher drawing"),
      length: readUint32(drawing, candidate + 4, "Invalid DOC Escher drawing"),
    });
    if (canRead(offset)) {
      const header = read(offset);
      if (isEscherType(header.type) && header.length <= end - offset - 8) return offset;
    }
    for (
      let candidate = offset + 1;
      candidate + 8 <= end && candidate - offset <= 16;
      candidate += 1
    ) {
      const header = read(candidate);
      if (isEscherType(header.type) && header.length <= end - candidate - 8) return candidate;
    }
    throw new DocParseError("Invalid DOC Escher drawing: no valid record boundary", {
      part: "data",
      path,
      offset,
      reason: "invalid-record",
    });
  };

  const walk = (start: number, end: number): void => {
    if (start + 8 > end) return;
    let offset = alignRecord(start, end, "Escher");
    while (offset + 8 <= end) {
      const version = readUint16(drawing, offset, "Invalid DOC Escher drawing") & 0x000f;
      const type = readUint16(drawing, offset + 2, "Invalid DOC Escher drawing");
      const length = readUint32(drawing, offset + 4, "Invalid DOC Escher drawing");
      const body = offset + 8;
      if (length > end - body) {
        throw new DocParseError("Invalid DOC Escher drawing: record is truncated", {
          part: "data",
          path: "Escher",
          offset,
          length,
          byteRange: [offset, offset + 8 + length],
          reason: "out-of-range",
        });
      }
      if (version === 0x0f) {
        if (type === 0xf004) parseShapeContainer(body, body + length);
        else walk(body, body + length);
      } else if (type === 0xf007) {
        pictures.push(parsePictureStoreEntry(body, body + length, dataStream));
      } else if (type >= 0xf01a && type <= 0xf01f) {
        const blip = parseEmbeddedBlip(drawing.subarray(body, body + length));
        if (!blip) {
          throw new DocParseError("Invalid DOC Escher drawing: unsupported BLIP record", {
            part: "data",
            path: "Escher/BLIP",
            offset,
            length,
            reason: "unsupported-required-structure",
          });
        }
        pictures.push(blip);
      }
      if (body + length >= end) break;
      offset = alignRecord(body + length, end, "Escher");
    }
  };

  const parseShapeContainer = (start: number, end: number): void => {
    if (start + 8 > end) return;
    let shapeId: number | undefined;
    let blipIndex: number | undefined;
    let cursor = alignRecord(start, end, "Escher/SpContainer");
    while (cursor + 8 <= end) {
      const type = readUint16(drawing, cursor + 2, "Invalid DOC Escher shape");
      const length = readUint32(drawing, cursor + 4, "Invalid DOC Escher shape");
      const body = cursor + 8;
      if (length > end - body) {
        throw new DocParseError("Invalid DOC Escher shape: child record is truncated", {
          part: "data",
          path: "Escher/SpContainer",
          offset: cursor,
          length,
          byteRange: [cursor, cursor + 8 + length],
          reason: "out-of-range",
        });
      }
      if (type === 0xf00a) shapeId = readUint32(drawing, body, "Invalid DOC Escher shape id");
      if (type === 0xf00b) blipIndex = parseBlipProperty(cursor, cursor + 8 + length);
      if (body + length >= end) break;
      cursor = alignRecord(body + length, end, "Escher/SpContainer");
    }
    if (shapeId !== undefined && blipIndex !== undefined) {
      const picture = pictures[blipIndex - 1];
      if (!picture || picture.type === "unknown") {
        throw new DocParseError("Invalid DOC Escher shape: BLIP reference is out of range", {
          part: "data",
          path: "Escher/OPT/BLIPTODISPLAY",
          offset: start,
          reason: "out-of-range",
        });
      }
      picturesByShapeId.set(shapeId, picture);
    }
  };

  const parseBlipProperty = (offset: number, end: number): number | undefined => {
    requireRange(drawing, offset, 6, "Invalid DOC Escher properties", {
      part: "data",
      path: "Escher/OPT",
      offset,
      length: 6,
      reason: "out-of-range",
    });
    const count = (readUint16(drawing, offset, "Invalid DOC Escher properties") >> 4) & 0x0fff;
    if (offset + count * 6 > end) {
      throw new DocParseError("Invalid DOC Escher properties: property table is truncated", {
        part: "data",
        path: "Escher/OPT",
        offset,
        length: count * 6,
        reason: "out-of-range",
      });
    }
    let complexBytes = 0;
    for (let index = 0; index < count; index += 1) {
      const id = readUint16(drawing, offset + 8 + index * 6, "Invalid DOC Escher property");
      const value = readUint32(drawing, offset + 8 + index * 6 + 2, "Invalid DOC Escher property");
      if ((id & 0x8000) !== 0) complexBytes += value;
      if ((id & 0x3fff) === 0x0104 && (id & 0x4000) !== 0) return value;
    }
    requireRange(drawing, end, complexBytes, "Invalid DOC Escher properties", {
      part: "data",
      path: "Escher/OPT",
      offset: end,
      length: complexBytes,
      reason: "out-of-range",
    });
    return undefined;
  };

  const parsePictureStoreEntry = (
    offset: number,
    end: number,
    data: Uint8Array,
  ): LegacyPictureData => {
    requireRange(drawing, offset, 36, "Invalid DOC Escher BSE", {
      part: "data",
      path: "Escher/BSE",
      offset,
      length: 36,
      reason: "out-of-range",
    });
    const delayOffset = readUint32(drawing, offset + 28, "Invalid DOC Escher picture");
    const nameLength = drawing[offset + 33] ?? 0;
    const embedded = offset + 36 + nameLength;
    let embeddedError: unknown;
    let embeddedPicture: LegacyPictureData | undefined;
    if (embedded + 8 <= end) {
      try {
        embeddedPicture = parseEmbeddedBlip(drawing.subarray(embedded, end));
      } catch (error) {
        embeddedError = error;
      }
    }
    if (embeddedPicture) return embeddedPicture;
    if (delayOffset === 0xffffffff) {
      throw (
        embeddedError ??
        new DocParseError("Invalid DOC Escher BSE: embedded BLIP is truncated", {
          part: "data",
          path: "Escher/BSE/BLIP",
          offset: embedded,
          reason: "out-of-range",
        })
      );
    }
    return parseBlip(data, delayOffset);
  };

  walk(0, drawing.byteLength);
  return new Map([...picturesByShapeId].filter(([, picture]) => picture.data.byteLength > 0));
}

/** Parse PlcspaMom into main-document CP → shape-id anchors. */
export function parseBodyPictureAnchors(table: Uint8Array, range: NumberPair): Map<number, number> {
  const message = "Invalid DOC picture anchor table: it is outside the table stream";
  requireRange(table, range.offset, range.length, message, {
    part: "table",
    path: "PlcspaMom",
    offset: range.offset,
    length: range.length,
  });
  const entryLength = 4 + 26;
  if (range.length < entryLength || (range.length - 4) % entryLength !== 0) {
    throw new DocParseError("Invalid DOC picture anchor table: malformed CP count", {
      part: "table",
      path: "PlcspaMom",
      offset: range.offset,
      length: range.length,
      byteRange: [range.offset, range.offset + range.length],
      reason: "invalid-record-length",
    });
  }
  const count = (range.length - 4) / entryLength;
  const anchors = new Map<number, number>();
  for (let index = 0; index < count; index += 1) {
    const cp = readUint32(table, range.offset + index * 4, message);
    const entryOffset = range.offset + (count + 1) * 4 + index * 26;
    const shapeId = readUint32(table, entryOffset, message);
    anchors.set(cp, shapeId);
  }
  return anchors;
}

function imageTypeFromBytes(data: Uint8Array): LegacyPictureData["type"] | undefined {
  if (data[0] === 0x89 && data[1] === 0x50 && data[2] === 0x4e && data[3] === 0x47) return "png";
  if (data[0] === 0xff && data[1] === 0xd8) return "jpeg";
  if (data[0] === 0x47 && data[1] === 0x49 && data[2] === 0x46) return "gif";
  if (data[0] === 0x42 && data[1] === 0x4d) return "bmp";
  if ((data[0] === 0x49 && data[1] === 0x20) || (data[0] === 0x4d && data[1] === 0x4d))
    return "tiff";
  return undefined;
}

function parseEmbeddedBlip(bytes: Uint8Array): LegacyPictureData | undefined {
  const requireBlipRange = (offset: number, length: number): void => {
    requireRange(bytes, offset, length, "Invalid DOC Escher BLIP: record is truncated", {
      part: "data",
      path: "Escher/BLIP",
      offset,
      length,
      byteRange: [offset, offset + length],
      reason: "out-of-range",
    });
  };
  requireBlipRange(0, 4);
  const header = readUint16(bytes, 0, "Invalid DOC Escher BLIP");
  const recordType = readUint16(bytes, 2, "Invalid DOC Escher BLIP");
  if (recordType < 0xf01a || recordType > 0xf01f) {
    return undefined;
  }
  const instance = header >> 4;
  const uidCount = instance % 2 === 0 ? 1 : 2;
  const metafileHeaderLength = recordType <= 0xf01c ? 34 : 0;
  const start = 8 + uidCount * 16 + metafileHeaderLength;
  requireBlipRange(start, 1);
  const type =
    recordType === 0xf01a
      ? "emf"
      : recordType === 0xf01b
        ? "wmf"
        : (imageTypeFromBytes(bytes.subarray(start)) ?? "unknown");
  const data = bytes.subarray(start);
  const dimensions = readImageDimensions(data, type);
  return { type, data, ...dimensions };
}

function parseBlip(dataStream: Uint8Array, reference: number): LegacyPictureData {
  if (reference >= dataStream.byteLength) {
    throw new DocParseError("Invalid DOC Data stream: BLIP reference is out of range", {
      part: "data",
      path: "Data",
      recordName: "BLIP",
      offset: reference,
      length: dataStream.byteLength,
      reason: "out-of-range",
    });
  }
  const picture = parseEmbeddedBlip(dataStream.subarray(reference));
  if (!picture) {
    throw new DocParseError("Invalid DOC Data stream: unsupported BLIP record", {
      part: "data",
      path: "Data",
      recordName: "BLIP",
      offset: reference,
      reason: "unsupported-record",
    });
  }
  return picture;
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
