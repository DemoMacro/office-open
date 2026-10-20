import type {
  BulletCharOptions,
  TextCharacterPropertiesOptions,
  TextParagraphPropertiesOptions,
  SchemeColorOptions,
} from "@office-open/core";

import { PptParseError } from "../errors";
import { RecordType, type RecordNode } from "../records";
import type { LegacyTextStyle, LegacyTextStyleCollection } from "./models";

const PARAGRAPH_FLAGS_MASK = 0x0000000f;
const BULLET_CHARACTER_MASK = 0x00000080;
const BULLET_FONT_MASK = 0x00000010;
const BULLET_SIZE_MASK = 0x00000040;
const BULLET_COLOR_MASK = 0x00000020;
const PARAGRAPH_ALIGNMENT_MASK = 0x00000800;
const LINE_SPACING_MASK = 0x00001000;
const SPACE_BEFORE_MASK = 0x00002000;
const SPACE_AFTER_MASK = 0x00004000;
const TEXT_OFFSET_MASK = 0x00000100;
const BULLET_OFFSET_MASK = 0x00000400;
const DEFAULT_TAB_SIZE_MASK = 0x00008000;
const TAB_STOPS_MASK = 0x00100000;
const FONT_ALIGNMENT_MASK = 0x00010000;
const WRAP_FLAGS_MASK = 0x000e0000;
const TEXT_DIRECTION_MASK = 0x00200000;
const BULLET_BLIP_MASK = 0x01000000;
const BULLET_SCHEME_MASK = 0x02000000;
const BULLET_HAS_SCHEME_MASK = 0x04000000;
const PARAGRAPH_EXTENSION_MASKS = [
  0x00800000, 0x08000000, 0x10000000, 0x20000000, 0x40000000, 0x80000000,
] as const;
const PP10_CHARACTER_EXTENSION_MASK = 0x00100000;
const NEW_EA_FONT_MASK = 0x01000000;
const CS_FONT_MASK = 0x02000000;
const PP11_CHARACTER_EXTENSION_MASK = 0x04000000;
const CHARACTER_RESERVED_MASKS = [
  0x08000000, 0x10000000, 0x20000000, 0x40000000, 0x80000000,
] as const;
const CHARACTER_FLAGS_MASK = 0x0000ffff;
const FONT_INDEX_MASK = 0x00010000;
const ASIAN_FONT_INDEX_MASK = 0x00200000;
const ANSI_FONT_INDEX_MASK = 0x00400000;
const SYMBOL_FONT_INDEX_MASK = 0x00800000;
const FONT_SIZE_MASK = 0x00020000;
const FONT_COLOR_MASK = 0x00040000;
const BASELINE_MASK = 0x00080000;

const ALIGNMENTS = [
  "left",
  "center",
  "right",
  "justify",
  "distribute",
  "thaiDistributed",
  "lowJustification",
] as const;

interface StyleReader {
  readonly view: DataView;
  readonly start: number;
  readonly end: number;
}

interface ParsedCollection {
  readonly collection: LegacyTextStyleCollection;
  readonly next: number;
}

function styleReader(view: DataView, record: RecordNode): StyleReader {
  if (record.type !== RecordType.styleTextPropAtom || record.length < 10) {
    throw new PptParseError("Corrupt legacy PowerPoint text style atom", {
      part: "record",
      recordType: record.type,
      recordName: "StyleTextPropAtom",
      offset: record.offset,
      length: record.length,
      reason: "invalid-text-style-record",
    });
  }
  return { view, start: record.offset + 8, end: record.offset + 8 + record.length };
}

function requireBytes(reader: StyleReader, offset: number, length: number, reason: string): void {
  if (offset < reader.start || offset + length > reader.end) {
    throw new PptParseError("Corrupt legacy PowerPoint text style atom: truncated field", {
      part: "record",
      recordType: RecordType.styleTextPropAtom,
      recordName: "StyleTextPropAtom",
      offset,
      length,
      byteRange: [offset, offset + length],
      reason,
    });
  }
}

function readInt16(reader: StyleReader, offset: number): number {
  requireBytes(reader, offset, 2, "truncated-int16");
  return reader.view.getInt16(offset, true);
}

function readInt32(reader: StyleReader, offset: number): number {
  requireBytes(reader, offset, 4, "truncated-int32");
  return reader.view.getInt32(offset, true);
}

function readCoverage(reader: StyleReader, offset: number): number {
  const characters = reader.view.getUint16(offset, true);
  if (characters > 0xffff) {
    throw new PptParseError("Corrupt legacy PowerPoint text style atom: invalid coverage", {
      part: "record",
      recordType: RecordType.styleTextPropAtom,
      recordName: "StyleTextPropAtom",
      offset,
      reason: "invalid-style-coverage",
    });
  }
  return characters;
}

function alignment(value: number): TextParagraphPropertiesOptions["alignment"] {
  const result = ALIGNMENTS[value];
  if (!result) {
    throw new PptParseError("Corrupt legacy PowerPoint text style atom: unknown alignment", {
      part: "record",
      recordType: RecordType.styleTextPropAtom,
      recordName: "StyleTextPropAtom",
      reason: "unknown-text-alignment",
    });
  }
  return result;
}

function fontColor(
  value: number,
): string | { type: "solid"; color: SchemeColorOptions } | undefined {
  const colorIndex = (value >>> 24) & 0xff;
  if (colorIndex === 0xff) return undefined;
  if (colorIndex !== 0xfe) {
    const schemeColors = [
      "bg1",
      "tx1",
      "bg2",
      "tx2",
      "accent1",
      "accent2",
      "accent3",
      "accent4",
    ] as const;
    const scheme = schemeColors[colorIndex];
    if (!scheme) {
      throw new PptParseError(
        "Corrupt legacy PowerPoint text style atom: invalid font color index",
        {
          part: "record",
          recordType: RecordType.styleTextPropAtom,
          recordName: "StyleTextPropAtom",
          reason: "invalid-font-color-index",
        },
      );
    }
    return { type: "solid", color: { value: scheme } };
  }
  const blue = value & 0xff;
  const green = (value >>> 8) & 0xff;
  const red = (value >>> 16) & 0xff;
  return [red, green, blue].map((channel) => channel.toString(16).padStart(2, "0")).join("");
}

function parseParagraph(reader: StyleReader, offset: number): ParsedCollection {
  const characters = readCoverage(reader, offset);
  let cursor = offset + 2;
  const indentLevel = readInt16(reader, cursor);
  cursor += 2;
  const mask = readInt32(reader, cursor);
  cursor += 4;
  const properties: TextParagraphPropertiesOptions = {};

  if ((mask & PARAGRAPH_FLAGS_MASK) !== 0) cursor += 2;
  if ((mask & BULLET_CHARACTER_MASK) !== 0) {
    const charCode = readInt16(reader, cursor);
    properties.bullet = { type: "char", char: String.fromCharCode(charCode) };
    cursor += 2;
  }
  if ((mask & BULLET_FONT_MASK) !== 0) cursor += 2;
  if ((mask & BULLET_SIZE_MASK) !== 0) {
    const size = readInt16(reader, cursor);
    const bullet: BulletCharOptions =
      properties.bullet?.type === "char" ? properties.bullet : { type: "char" };
    properties.bullet = { ...bullet, size: Math.round(size / 10) };
    cursor += 2;
  }
  if ((mask & BULLET_COLOR_MASK) !== 0) {
    const color = fontColor(readInt32(reader, cursor));
    const bulletColor =
      typeof color === "string" ? color : (color?.color as BulletCharOptions["color"]);
    const bullet: BulletCharOptions =
      properties.bullet?.type === "char" ? properties.bullet : { type: "char" };
    properties.bullet = { ...bullet, ...(bulletColor ? { color: bulletColor } : {}) };
    cursor += 4;
  }
  if ((mask & PARAGRAPH_ALIGNMENT_MASK) !== 0) {
    properties.alignment = alignment(readInt16(reader, cursor));
    cursor += 2;
  }
  if ((mask & LINE_SPACING_MASK) !== 0) {
    const spacing = readInt16(reader, cursor);
    if (spacing < 0) properties.lineSpacingPercent = Math.round((spacing * 100) / -100_000);
    else properties.lineSpacingPoints = spacing / 100;
    cursor += 2;
  }
  if ((mask & SPACE_BEFORE_MASK) !== 0) {
    properties.spaceBefore = readInt16(reader, cursor) / 100;
    cursor += 2;
  }
  if ((mask & SPACE_AFTER_MASK) !== 0) {
    properties.spaceAfter = readInt16(reader, cursor) / 100;
    cursor += 2;
  }
  if ((mask & TEXT_OFFSET_MASK) !== 0) {
    properties.marginIndent = readInt16(reader, cursor) * 635;
    cursor += 2;
  }
  if ((mask & BULLET_OFFSET_MASK) !== 0) {
    properties.indent = readInt16(reader, cursor) * 635;
    cursor += 2;
  }
  if ((mask & DEFAULT_TAB_SIZE_MASK) !== 0) {
    properties.defTabSize = readInt16(reader, cursor) * 635;
    cursor += 2;
  }
  if ((mask & TAB_STOPS_MASK) !== 0) {
    const count = readInt16(reader, cursor);
    cursor += 2;
    properties.tabStops = Array.from({ length: count }, (_, index) => ({
      position: readInt16(reader, cursor + index * 4) * 635,
      alignment: "left" as const,
    }));
    cursor += count * 4;
  }
  if ((mask & FONT_ALIGNMENT_MASK) !== 0) {
    const alignments = ["base", "top", "center", "bottom"] as const;
    properties.fontAlignment = alignments[readInt16(reader, cursor)] ?? "auto";
    cursor += 2;
  }
  if ((mask & WRAP_FLAGS_MASK) !== 0) cursor += 2;
  if ((mask & TEXT_DIRECTION_MASK) !== 0) cursor += 2;
  if ((mask & 0x00400000) !== 0) {
    throw new PptParseError("Corrupt legacy PowerPoint text style atom: reserved mask", {
      part: "record",
      recordType: RecordType.styleTextPropAtom,
      recordName: "StyleTextPropAtom",
      reason: "reserved-text-style-mask",
    });
  }
  if ((mask & BULLET_BLIP_MASK) !== 0) cursor += 2;
  if ((mask & BULLET_SCHEME_MASK) !== 0) cursor += 4;
  if ((mask & BULLET_HAS_SCHEME_MASK) !== 0) cursor += 2;
  for (const extensionMask of PARAGRAPH_EXTENSION_MASKS) {
    if ((mask & extensionMask) !== 0) cursor += 0;
  }
  if (indentLevel > 0) properties.indentLevel = indentLevel;

  return {
    collection: {
      characters,
      paragraph: Object.keys(properties).length > 0 ? properties : undefined,
    },
    next: cursor,
  };
}

function parseCharacter(
  reader: StyleReader,
  offset: number,
  fonts: ReadonlyMap<number, string>,
): ParsedCollection {
  const characters = readCoverage(reader, offset);
  let cursor = offset + 2;
  cursor += 2;
  const mask = readInt32(reader, cursor);
  cursor += 4;
  const properties: TextCharacterPropertiesOptions = {};

  if ((mask & PP10_CHARACTER_EXTENSION_MASK) !== 0) cursor += 4;
  if ((mask & NEW_EA_FONT_MASK) !== 0) cursor += 2;
  if ((mask & CS_FONT_MASK) !== 0) cursor += 2;
  if ((mask & PP11_CHARACTER_EXTENSION_MASK) !== 0) cursor += 4;
  for (const extensionMask of CHARACTER_RESERVED_MASKS) {
    if ((mask & extensionMask) !== 0) cursor += 0;
  }
  if ((mask & CHARACTER_FLAGS_MASK) !== 0) {
    const flags = readInt16(reader, cursor);
    cursor += 2;
    if ((flags & 0x0001) !== 0) properties.bold = true;
    if ((flags & 0x0002) !== 0) properties.italic = true;
    if ((flags & 0x0004) !== 0) properties.underline = "single";
    if ((flags & 0x0010) !== 0) properties.shadow = true;
    if ((flags & 0x0100) !== 0) properties.strike = "singleStrike";
  }
  type CharacterScript = "latin" | "eastAsia" | "complexScript" | "symbol";
  const fontIndexes: Partial<Record<CharacterScript, string>> = {};
  for (const [fontMask, script] of [
    [FONT_INDEX_MASK, "latin"],
    [ASIAN_FONT_INDEX_MASK, "eastAsia"],
    [ANSI_FONT_INDEX_MASK, "complexScript"],
    [SYMBOL_FONT_INDEX_MASK, "symbol"],
  ] as const) {
    if ((mask & fontMask) === 0) continue;
    const index = readInt16(reader, cursor);
    cursor += 2;
    const fontName = fonts.get(index);
    if (fontName !== undefined) fontIndexes[script] = fontName;
  }
  if ((mask & FONT_SIZE_MASK) !== 0) {
    const size = readInt16(reader, cursor);
    cursor += 2;
    if (size <= 0 || size > 4000) {
      throw new PptParseError("Corrupt legacy PowerPoint text style atom: invalid font size", {
        part: "record",
        recordType: RecordType.styleTextPropAtom,
        recordName: "StyleTextPropAtom",
        offset: cursor - 2,
        reason: "invalid-font-size",
      });
    }
    properties.size = size;
  }
  if ((mask & FONT_COLOR_MASK) !== 0) {
    const fill = fontColor(readInt32(reader, cursor));
    if (fill) properties.fill = fill;
    cursor += 4;
  }
  if ((mask & BASELINE_MASK) !== 0) {
    properties.baseline = readInt16(reader, cursor) / 1000;
    cursor += 2;
  }
  const fontNames = Object.values(fontIndexes);
  if (fontNames.length === 1) {
    properties.font = fontNames[0];
  } else if (fontNames.length > 0) {
    properties.font = fontIndexes;
  }

  const knownMask =
    PP10_CHARACTER_EXTENSION_MASK |
    NEW_EA_FONT_MASK |
    CS_FONT_MASK |
    PP11_CHARACTER_EXTENSION_MASK |
    CHARACTER_RESERVED_MASKS.reduce((left, right) => left | right, 0) |
    CHARACTER_FLAGS_MASK |
    FONT_INDEX_MASK |
    ASIAN_FONT_INDEX_MASK |
    ANSI_FONT_INDEX_MASK |
    SYMBOL_FONT_INDEX_MASK |
    FONT_SIZE_MASK |
    FONT_COLOR_MASK |
    BASELINE_MASK;
  if ((mask & ~knownMask) !== 0) {
    throw new PptParseError(
      `Corrupt legacy PowerPoint text style atom: unsupported character mask 0x${(mask >>> 0).toString(16)}`,
      {
        part: "record",
        recordType: RecordType.styleTextPropAtom,
        recordName: "StyleTextPropAtom",
        reason: "unsupported-text-style-mask",
      },
    );
  }
  return {
    collection: {
      characters,
      character: Object.keys(properties).length > 0 ? properties : undefined,
    },
    next: cursor,
  };
}

function parseCollections(
  reader: StyleReader,
  parser: (reader: StyleReader, offset: number) => ParsedCollection,
  textLength: number,
): { collections: LegacyTextStyleCollection[]; end: number } {
  const collections: LegacyTextStyleCollection[] = [];
  let cursor = reader.start;
  let covered = 0;
  while (covered < textLength) {
    if (cursor >= reader.end) {
      throw new PptParseError("Corrupt legacy PowerPoint text style atom: truncated styles", {
        part: "record",
        recordType: RecordType.styleTextPropAtom,
        recordName: "StyleTextPropAtom",
        offset: cursor,
        reason: "truncated-style-collection",
      });
    }
    const parsed = parser(reader, cursor);
    const collection =
      parsed.collection.characters === 0 && parsed.next >= reader.end
        ? { ...parsed.collection, characters: textLength - covered }
        : parsed.collection;
    collections.push(collection);
    covered += collection.characters;
    cursor = parsed.next;
  }
  if (covered > textLength + 1 || cursor > reader.end) {
    throw new PptParseError(
      `Corrupt legacy PowerPoint text style atom: style coverage mismatch (${covered} covered, ${textLength} expected)`,
      {
        part: "record",
        recordType: RecordType.styleTextPropAtom,
        recordName: "StyleTextPropAtom",
        offset: reader.start,
        length: reader.end - reader.start,
        reason: "invalid-style-coverage-total",
      },
    );
  }
  return { collections, end: cursor };
}

export function parseTextStyle(
  view: DataView,
  record: RecordNode,
  textLength: number,
  fonts: ReadonlyMap<number, string>,
): LegacyTextStyle {
  const reader = styleReader(view, record);
  const paragraphs = parseCollections(reader, parseParagraph, textLength);
  if (paragraphs.end >= reader.end) {
    throw new PptParseError("Corrupt legacy PowerPoint text style atom: missing character styles", {
      part: "record",
      recordType: RecordType.styleTextPropAtom,
      recordName: "StyleTextPropAtom",
      offset: paragraphs.end,
      reason: "missing-character-style",
    });
  }
  const characterReader: StyleReader = { ...reader, start: paragraphs.end };
  const characters = parseCollections(
    characterReader,
    (styleReader, offset) => parseCharacter(styleReader, offset, fonts),
    textLength,
  );
  return { paragraphs: paragraphs.collections, characters: characters.collections };
}
