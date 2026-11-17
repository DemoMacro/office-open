import type {
  AlignmentOptions,
  BorderOptions,
  BorderSideOptions,
  CellFillOptions,
  CellStyleCategory,
  CellStyleXfOptions,
  ColorOptions,
  FontOptions,
} from "@office-open/xlsx";

import { XlsParseError } from "../errors";
import type { BiffRecord, BiffVersion } from "../records";
import { decodeBiffString, readBiff5String, readBiff8String } from "./biff";
import type { XfStyle } from "./models";

export interface LegacyStyleRecord {
  readonly xfIndex: number;
  readonly builtIn: boolean;
  readonly builtInData?: number;
  readonly builtinId?: number;
  readonly iLevel?: number;
  readonly name?: string;
}

export interface LegacyStyleExtension {
  readonly record: BiffRecord;
  readonly styleRecordIndex?: number;
  readonly name: string;
  readonly builtIn: boolean;
  readonly hidden: boolean;
  readonly customBuiltin: boolean;
  readonly category: CellStyleCategory;
  readonly builtInData: number;
  readonly builtinId?: number;
  readonly style: CellStyleXfOptions;
}
export function parseFont(
  body: Uint8Array,
  isBiff8: boolean,
  encoding: string,
): FontOptions | undefined {
  if (body.byteLength < 15)
    throw new XlsParseError("Invalid legacy XLS file: truncated FONT record");
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  // MS-XLS 2.5.146: Font.name is a shortXLUnicodeString — a 1-byte cch (not
  // the 2-byte XLUnicodeString count), preceded by 14 fixed bytes.
  let fontName = "";
  const characterCount = body[14]!;
  fontName = isBiff8
    ? decodeBiffString(body, view, 16, characterCount, body[15] ?? 0, encoding).value
    : new TextDecoder(encoding).decode(body.subarray(15, 15 + characterCount));
  const weight = view.getUint16(6, true);
  return {
    size: view.getUint16(0, true) / 20,
    bold: weight >= 700,
    italic: (view.getUint16(2, true) & 0x0002) !== 0,
    color: { indexed: view.getUint16(4, true) },
    underline: body[10] === 1,
    font: fontName,
  };
}

export function parseFormat(
  body: Uint8Array,
  isBiff8: boolean,
  encoding: string,
): { id: number; code: string } | undefined {
  if (body.byteLength < 4)
    throw new XlsParseError("Invalid legacy XLS file: truncated FORMAT record");
  const id = new DataView(body.buffer, body.byteOffset, body.byteLength).getUint16(0, true);
  let code: string;
  code = (isBiff8 ? readBiff8String(body, 2, encoding) : readBiff5String(body, 2, encoding)).value;
  return {
    id,
    code,
  };
}

export function parsePalette(body: Uint8Array): string[] {
  if (body.byteLength < 2) return [];
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const count = view.getUint16(0, true);
  if (body.byteLength < 2 + count * 4) return [];
  return Array.from({ length: count }, (_, index) => {
    const value = view.getUint32(2 + index * 4, true);
    const blue = (value & 0xff).toString(16).padStart(2, "0");
    const green = ((value >> 8) & 0xff).toString(16).padStart(2, "0");
    const red = ((value >> 16) & 0xff).toString(16).padStart(2, "0");
    return `FF${red}${green}${blue}`;
  });
}

const BORDER_STYLES: readonly BorderOptions["style"][] = [
  "none",
  "thin",
  "medium",
  "dashed",
  "dotted",
  "thick",
  "double",
  "hair",
  "mediumDashed",
  "dashDot",
  "mediumDashDot",
  "dashDotDot",
  "mediumDashDotDot",
  "slantDashDot",
];

const PATTERN_TYPES: readonly NonNullable<CellFillOptions["patternType"]>[] = [
  "none",
  "solid",
  "mediumGray",
  "darkGray",
  "lightGray",
  "darkHorizontal",
  "darkVertical",
  "darkDown",
  "darkUp",
  "darkGrid",
  "darkTrellis",
  "lightHorizontal",
  "lightVertical",
  "lightDown",
  "lightUp",
  "lightGrid",
  "lightTrellis",
  "gray125",
  "gray0625",
];

const HORIZONTAL_ALIGNMENT: readonly AlignmentOptions["horizontal"][] = [
  "general",
  "left",
  "center",
  "right",
  "fill",
  "justify",
  "centerContinuous",
  "distributed",
];

const VERTICAL_ALIGNMENT: readonly AlignmentOptions["vertical"][] = [
  "top",
  "center",
  "bottom",
  "justify",
  "distributed",
];

export function parseXf(body: Uint8Array, version: BiffVersion, palette: string[]): XfStyle {
  const expectedLength = version === 8 ? 20 : 16;
  if (body.byteLength < expectedLength)
    throw new XlsParseError("Invalid legacy XLS file: truncated extended format", {
      part: "stream",
      path: "Workbook/globals/XF",
      recordType: 0x00e0,
      recordName: "XF",
      reason: "truncated-record",
      length: expectedLength,
    });
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const cellOptions = view.getUint16(4, true);
  const isStyle = (cellOptions & 0x0004) !== 0;
  const style: XfStyle = {
    fontIndex: view.getUint16(0, true),
    numberFormatId: view.getUint16(2, true),
    isStyle,
    parentIndex: cellOptions >> 4,
    usedAttributes: parseUsedAttributes(version === 8 ? view.getUint8(9) : view.getUint8(7)),
  };
  const alignment = parseAlignment(body, view, version);
  if (alignment) style.alignment = alignment;
  style.border = parseBorder(view, version, palette);
  style.fill = parseFill(view, version, palette);
  if ((cellOptions & 0x0001) === 0 || (cellOptions & 0x0002) !== 0)
    style.protection = {
      locked: (cellOptions & 0x0001) !== 0,
      hidden: (cellOptions & 0x0002) !== 0,
    };
  if (version === 8 && (cellOptions & 0x0008) !== 0) style.quotePrefix = true;
  const useOwn = (used: boolean) => (isStyle ? !used : used);
  if (!useOwn(style.usedAttributes.alignment)) style.alignment = undefined;
  if (!useOwn(style.usedAttributes.border)) style.border = undefined;
  if (!useOwn(style.usedAttributes.fill)) style.fill = undefined;
  if (!useOwn(style.usedAttributes.protection)) style.protection = undefined;
  return style;
}

export function parseStyle(
  body: Uint8Array,
  isBiff8: boolean,
  encoding: string,
  record: BiffRecord,
): LegacyStyleRecord {
  if (body.byteLength < 2) throw invalidStyle(record, "truncated-header");
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const value = view.getUint16(0, true);
  const builtIn = (value & 0x8000) !== 0;
  if (builtIn) {
    if (body.byteLength !== 4) throw invalidStyle(record, "invalid-built-in-data-size");
    const builtinId = body[2]!;
    const iLevel = body[3]!;
    if (builtinId > 0x09) throw invalidStyle(record, `invalid-built-in-id-${builtinId}`);
    return {
      xfIndex: value & 0x0fff,
      builtIn,
      builtInData: view.getUint16(2, true),
      builtinId,
      ...(builtinId === 0x01 || builtinId === 0x02 ? { iLevel } : {}),
    };
  }
  const decoded = isBiff8 ? readBiff8String(body, 2, encoding) : readBiff5String(body, 2, encoding);
  if (decoded.offset !== body.byteLength) throw invalidStyle(record, "trailing-style-data");
  const name = decoded.value;
  if (name.length === 0 || name.length > 255) throw invalidStyle(record, "invalid-style-name");
  return { xfIndex: value & 0x0fff, builtIn, name };
}

export function parseStyleExt(
  body: Uint8Array,
  palette: string[],
  record: BiffRecord,
): LegacyStyleExtension {
  if (body.byteLength < 20) throw invalidStyle(record, "truncated-header");
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  if (view.getUint16(0, true) !== 0x0892) throw invalidStyle(record, "invalid-record-type");
  const flags = body[12]!;
  const builtIn = (flags & 0x01) !== 0;
  const hidden = (flags & 0x02) !== 0;
  const customBuiltin = (flags & 0x04) !== 0;
  if ((flags & 0xf8) !== 0) throw invalidStyle(record, "invalid-style-flags");
  if (customBuiltin && !builtIn) throw invalidStyle(record, "custom-style-not-built-in");
  const category = STYLE_CATEGORIES[body[13]!];
  if (category === undefined) throw invalidStyle(record, `invalid-style-category-${body[13]}`);
  const builtinData = view.getUint16(14, true);
  if (!builtIn && builtinData !== 0xffff)
    throw invalidStyle(record, "invalid-custom-style-built-in-data");
  const nameValue = readLpWideString(body, 16, record, "style-name");
  if (nameValue.value.length === 0 || nameValue.value.length > 255)
    throw invalidStyle(record, "invalid-style-name");
  let offset = nameValue.offset;
  if (offset + 4 > body.byteLength) throw invalidStyle(record, "truncated-properties-header");
  const reserved = view.getUint16(offset, true);
  const propertyCount = view.getUint16(offset + 2, true);
  if (reserved !== 0) throw invalidStyle(record, "invalid-properties-reserved");
  offset += 4;
  const style: CellStyleXfOptions = {};
  for (let index = 0; index < propertyCount; index++) {
    if (offset + 4 > body.byteLength) throw invalidStyle(record, "truncated-property-header");
    const propertyType = view.getUint16(offset, true);
    const propertySize = view.getUint16(offset + 2, true);
    if (propertySize < 4 || offset + propertySize > body.byteLength)
      throw invalidStyle(record, `invalid-property-size-${propertySize}`);
    applyStyleProperty(
      style,
      propertyType,
      body.subarray(offset + 4, offset + propertySize),
      palette,
      record,
    );
    offset += propertySize;
  }
  if (offset !== body.byteLength) throw invalidStyle(record, "trailing-property-data");
  return {
    record,
    name: nameValue.value,
    builtIn,
    hidden,
    customBuiltin,
    category,
    builtInData: builtinData,
    ...(builtIn ? { builtinId: builtinData & 0xff } : {}),
    style,
  };
}

const STYLE_CATEGORIES: readonly CellStyleCategory[] = [
  "custom",
  "goodBadNeutral",
  "dataModel",
  "titleAndHeading",
  "themedCell",
  "numberFormat",
];

export function styleXfOptions(
  style: XfStyle,
  fonts: readonly FontOptions[],
  numberFormats: ReadonlyMap<number, string>,
): CellStyleXfOptions {
  const result: CellStyleXfOptions = {};
  const font = fonts[style.fontIndex];
  if (font) result.font = style.font ? { ...font, ...style.font } : font;
  else if (style.font) result.font = { ...DEFAULT_STYLE_FONT, ...style.font };
  const numFmt = numberFormats.get(style.numberFormatId);
  if (numFmt !== undefined) result.numFmt = numFmt;
  else if (style.numberFormatId !== 0) result.numFmtId = style.numberFormatId;
  if (style.alignment) result.alignment = style.alignment;
  if (style.border) result.border = style.border;
  if (style.fill) result.fill = style.fill;
  if (style.protection) result.protection = style.protection;
  if (style.quotePrefix) result.quotePrefix = true;
  return result;
}

const DEFAULT_STYLE_FONT: FontOptions = { font: "Arial", size: 10 };

const STYLE_PATTERN_TYPES: readonly NonNullable<CellFillOptions["patternType"]>[] = [
  "none",
  "solid",
  "mediumGray",
  "darkGray",
  "lightGray",
  "darkHorizontal",
  "darkVertical",
  "darkDown",
  "darkUp",
  "darkGrid",
  "darkTrellis",
  "lightHorizontal",
  "lightVertical",
  "lightDown",
  "lightUp",
  "lightGrid",
  "lightTrellis",
  "gray125",
  "gray0625",
];

function applyStyleProperty(
  style: CellStyleXfOptions,
  propertyType: number,
  data: Uint8Array,
  palette: string[],
  record: BiffRecord,
): void {
  const font = style.font ?? {};
  const alignment = style.alignment ?? {};
  const border = style.border ?? {};
  const fill = style.fill ?? {};
  const protection = style.protection ?? {};
  switch (propertyType) {
    case 0x00: {
      const pattern = readByte(data, record, propertyType);
      const patternType = STYLE_PATTERN_TYPES[pattern];
      if (patternType === undefined) throw invalidStyle(record, `invalid-fill-pattern-${pattern}`);
      style.fill = { ...fill, type: "pattern", patternType };
      return;
    }
    case 0x01:
      style.fill = { ...fill, foregroundColor: parseStyleColor(data, palette, record) };
      return;
    case 0x02:
      style.fill = { ...fill, backgroundColor: parseStyleColor(data, palette, record) };
      return;
    case 0x03:
      style.fill = parseStyleGradient(data, palette, record);
      return;
    case 0x04: {
      const gradient = style.fill;
      if (!gradient || gradient.type !== "gradient")
        throw invalidStyle(record, "gradient-stop-without-gradient");
      const stop = parseStyleGradientStop(data, palette, record);
      style.fill = { ...gradient, stops: [...(gradient.stops ?? []), stop] };
      return;
    }
    case 0x05:
      style.font = { ...font, color: parseStyleColor(data, palette, record) };
      return;
    case 0x06:
    case 0x07:
    case 0x08:
    case 0x09:
    case 0x0a:
    case 0x0b:
    case 0x0c: {
      const side = ["top", "bottom", "left", "right", "diagonal", "vertical", "horizontal"][
        propertyType - 0x06
      ] as keyof BorderSideOptions;
      const borderValue = parseStyleBorder(data, palette, record);
      style.border = { ...border, [side]: borderValue };
      return;
    }
    case 0x0d:
      style.border = { ...border, diagonalUp: readBoolean(data, record, propertyType) };
      return;
    case 0x0e:
      style.border = { ...border, diagonalDown: readBoolean(data, record, propertyType) };
      return;
    case 0x0f: {
      const value = readByte(data, record, propertyType);
      if (value !== 0xff && value > 7)
        throw invalidStyle(record, `invalid-horizontal-alignment-${value}`);
      if (value !== 0xff) alignment.horizontal = HORIZONTAL_ALIGNMENT[value];
      style.alignment = alignment;
      return;
    }
    case 0x10: {
      const value = readByte(data, record, propertyType);
      if (value !== 0xff && value > 4)
        throw invalidStyle(record, `invalid-vertical-alignment-${value}`);
      if (value !== 0xff) alignment.vertical = VERTICAL_ALIGNMENT[value];
      style.alignment = alignment;
      return;
    }
    case 0x11: {
      const value = readByte(data, record, propertyType);
      if (value > 0xb4 && value !== 0xff)
        throw invalidStyle(record, `invalid-text-rotation-${value}`);
      alignment.textRotation = value;
      style.alignment = alignment;
      return;
    }
    case 0x12: {
      const value = readUint16(data, record, propertyType);
      if (value > 15) throw invalidStyle(record, `invalid-indent-${value}`);
      alignment.indent = value;
      style.alignment = alignment;
      return;
    }
    case 0x13:
      alignment.readingOrder = readByte(data, record, propertyType);
      style.alignment = alignment;
      return;
    case 0x14:
      alignment.wrapText = readBoolean(data, record, propertyType);
      style.alignment = alignment;
      return;
    case 0x15:
      alignment.justifyLastLine = readBoolean(data, record, propertyType);
      style.alignment = alignment;
      return;
    case 0x16:
      alignment.shrinkToFit = readBoolean(data, record, propertyType);
      style.alignment = alignment;
      return;
    case 0x17:
      alignment.mergeCell = readBoolean(data, record, propertyType);
      style.alignment = alignment;
      return;
    case 0x18:
      style.font = { ...font, font: readLpWideString(data, 0, record, "font-name").value };
      return;
    case 0x19: {
      const value = readUint16(data, record, propertyType);
      if (value !== 0xffff && value !== 0x0190 && value !== 0x02bc)
        throw invalidStyle(record, `invalid-bold-${value}`);
      if (value !== 0xffff) style.font = { ...font, bold: value === 0x02bc };
      return;
    }
    case 0x1a: {
      const value = readByte(data, record, propertyType);
      const underline: FontOptions["underline"] =
        value === 1
          ? "single"
          : value === 2
            ? "double"
            : value === 3
              ? "singleAccounting"
              : value === 4
                ? "doubleAccounting"
                : undefined;
      if (value > 4) throw invalidStyle(record, `invalid-underline-${value}`);
      style.font = { ...font, ...(underline === undefined ? {} : { underline }) };
      return;
    }
    case 0x1b: {
      const value = readUint16(data, record, propertyType);
      if (value !== 0xffff && value > 2) throw invalidStyle(record, `invalid-script-${value}`);
      if (value !== 0xffff)
        style.font = {
          ...font,
          vertAlign: value === 1 ? "superscript" : value === 2 ? "subscript" : "baseline",
        };
      return;
    }
    case 0x1c:
      style.font = { ...font, italic: readBoolean(data, record, propertyType) };
      return;
    case 0x1d:
      style.font = { ...font, strike: readBoolean(data, record, propertyType) };
      return;
    case 0x1e:
      style.font = { ...font, outline: readBoolean(data, record, propertyType) };
      return;
    case 0x1f:
      style.font = { ...font, shadow: readBoolean(data, record, propertyType) };
      return;
    case 0x20:
      style.font = { ...font, condense: readBoolean(data, record, propertyType) };
      return;
    case 0x21:
      style.font = { ...font, extend: readBoolean(data, record, propertyType) };
      return;
    case 0x22:
      style.font = { ...font, charset: readByte(data, record, propertyType) };
      return;
    case 0x23: {
      const value = readByte(data, record, propertyType);
      if (value > 5) throw invalidStyle(record, `invalid-font-family-${value}`);
      style.font = { ...font, family: value };
      return;
    }
    case 0x24: {
      const value = readUint32(data, record, propertyType);
      if (value < 20 || value > 8191) throw invalidStyle(record, `invalid-font-size-${value}`);
      style.font = { ...font, size: value / 20 };
      return;
    }
    case 0x25:
      style.font = { ...font, scheme: parseFontScheme(data, record) };
      return;
    case 0x26:
      style.numFmt = readLpWideString(data, 0, record, "number-format").value;
      return;
    case 0x29:
      style.numFmtId = readUint16(data, record, propertyType);
      return;
    case 0x2a: {
      const value = readInt16(data, record, propertyType);
      if (value !== 255 && (value < -15 || value > 15))
        throw invalidStyle(record, `invalid-relative-indent-${value}`);
      if (value !== 255) alignment.relativeIndent = value;
      style.alignment = alignment;
      return;
    }
    case 0x2b:
      style.protection = { ...protection, locked: readBoolean(data, record, propertyType) };
      return;
    case 0x2c:
      style.protection = { ...protection, hidden: readBoolean(data, record, propertyType) };
      return;
    default:
      throw invalidStyle(record, `unsupported-property-${propertyType}`);
  }
}

function parseStyleColor(body: Uint8Array, palette: string[], record: BiffRecord): ColorOptions {
  if (body.byteLength !== 8) throw invalidStyle(record, `invalid-color-size-${body.byteLength}`);
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const flags = body[0]!;
  if ((flags & 0x01) === 0) throw invalidStyle(record, "invalid-color-rgba-flag");
  const type = flags >> 1;
  const icv = body[1]!;
  const tintShade = view.getInt16(2, true);
  if (tintShade === -32768) throw invalidStyle(record, "invalid-color-tint");
  const color: ColorOptions =
    type === 0
      ? { auto: true }
      : type === 1
        ? { indexed: icv }
        : type === 2
          ? { rgb: rgbaToArgb(view.getUint32(4, true)) }
          : type === 3
            ? { theme: icv }
            : type === 4
              ? {}
              : (() => {
                  throw invalidStyle(record, `unsupported-color-type-${type}`);
                })();
  if (tintShade !== 0 && type !== 0 && type !== 4) color.tint = tintShade / 32767;
  if (type === 1 && color.indexed !== undefined && color.indexed >= 8) {
    const customColor = palette[color.indexed - 8];
    if (customColor !== undefined) color.rgb = customColor as `FF${string}`;
  }
  return color;
}

function parseStyleBorder(
  body: Uint8Array,
  palette: string[],
  record: BiffRecord,
): NonNullable<BorderSideOptions["top"]> {
  if (body.byteLength !== 10) throw invalidStyle(record, `invalid-border-size-${body.byteLength}`);
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const styleValue = view.getUint16(8, true);
  const style = BORDER_STYLES[styleValue];
  if (style === undefined) throw invalidStyle(record, `invalid-border-style-${styleValue}`);
  return { style, color: parseStyleColor(body.subarray(0, 8), palette, record) };
}

function parseStyleGradient(
  body: Uint8Array,
  _palette: string[],
  record: BiffRecord,
): CellFillOptions {
  if (body.byteLength !== 44)
    throw invalidStyle(record, `invalid-gradient-size-${body.byteLength}`);
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const type = view.getUint32(0, true);
  if (type > 1) throw invalidStyle(record, `invalid-gradient-type-${type}`);
  const degree = view.getFloat64(4, true);
  const left = view.getFloat64(12, true);
  const right = view.getFloat64(20, true);
  const top = view.getFloat64(28, true);
  const bottom = view.getFloat64(36, true);
  if ([left, right, top, bottom].some((value) => value < 0 || value > 1))
    throw invalidStyle(record, "invalid-gradient-coordinate");
  return {
    type: "gradient",
    gradientType: type === 0 ? "linear" : "path",
    gradientDegree: degree,
    gradientLeft: left,
    gradientRight: right,
    gradientTop: top,
    gradientBottom: bottom,
    stops: [],
  };
}

function parseStyleGradientStop(
  body: Uint8Array,
  palette: string[],
  record: BiffRecord,
): NonNullable<CellFillOptions["stops"]>[number] {
  if (body.byteLength !== 18)
    throw invalidStyle(record, `invalid-gradient-stop-size-${body.byteLength}`);
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const position = view.getFloat64(2, true);
  if (position < 0 || position > 1) throw invalidStyle(record, "invalid-gradient-stop-position");
  return { position, color: parseStyleColor(body.subarray(10), palette, record) };
}

function readLpWideString(
  body: Uint8Array,
  offset: number,
  record: BiffRecord,
  field: string,
): { value: string; offset: number } {
  if (offset + 2 > body.byteLength) throw invalidStyle(record, `truncated-${field}`);
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const characterCount = view.getUint16(offset, true);
  const end = offset + 2 + characterCount * 2;
  if (end > body.byteLength) throw invalidStyle(record, `truncated-${field}`);
  return {
    value: new TextDecoder("utf-16le").decode(body.subarray(offset + 2, end)),
    offset: end,
  };
}

function readByte(body: Uint8Array, record: BiffRecord, propertyType: number): number {
  if (body.byteLength !== 1) throw invalidStyle(record, `invalid-property-size-${propertyType}`);
  return body[0]!;
}

function readBoolean(body: Uint8Array, record: BiffRecord, propertyType: number): boolean {
  const value = readByte(body, record, propertyType);
  if (value > 1) throw invalidStyle(record, `invalid-boolean-${propertyType}-${value}`);
  return value === 1;
}

function readUint16(body: Uint8Array, record: BiffRecord, propertyType: number): number {
  if (body.byteLength !== 2) throw invalidStyle(record, `invalid-property-size-${propertyType}`);
  return new DataView(body.buffer, body.byteOffset, body.byteLength).getUint16(0, true);
}

function readInt16(body: Uint8Array, record: BiffRecord, propertyType: number): number {
  if (body.byteLength !== 2) throw invalidStyle(record, `invalid-property-size-${propertyType}`);
  return new DataView(body.buffer, body.byteOffset, body.byteLength).getInt16(0, true);
}

function readUint32(body: Uint8Array, record: BiffRecord, propertyType: number): number {
  if (body.byteLength !== 4) throw invalidStyle(record, `invalid-property-size-${propertyType}`);
  return new DataView(body.buffer, body.byteOffset, body.byteLength).getUint32(0, true);
}

function invalidStyle(record: BiffRecord, reason: string): XlsParseError {
  return new XlsParseError(`Invalid legacy XLS file: ${reason}`, {
    part: "stream",
    path: "Workbook/globals/Style",
    recordType: record.code,
    recordName: record.code === 0x0892 ? "StyleExt" : "Style",
    offset: record.start,
    length: record.body.byteLength,
    reason,
  });
}

export function applyXfExt(
  record: BiffRecord,
  styleTable: Map<number, XfStyle>,
  palette: string[],
): void {
  const body = record.body;
  if (body.byteLength < 20) throw invalidXfExt(record, "truncated-header");
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  if (view.getUint16(0, true) !== 0x087d) throw invalidXfExt(record, "invalid-record-type");
  const index = view.getUint16(14, true);
  const style = styleTable.get(index);
  if (!style) throw invalidXfExt(record, `missing-XF-${index}`);
  const extensionCount = view.getUint16(18, true);
  let offset = 20;
  for (let extensionIndex = 0; extensionIndex < extensionCount; extensionIndex++) {
    if (offset + 4 > body.byteLength) throw invalidXfExt(record, "truncated-property-header");
    const extensionType = view.getUint16(offset, true);
    const length = view.getUint16(offset + 2, true);
    if (length < 4 || offset + length > body.byteLength)
      throw invalidXfExt(record, `invalid-property-length-${length}`);
    const data = body.subarray(offset + 4, offset + length);
    applyXfExtension(style, extensionType, data, palette, record);
    offset += length;
  }
  if (offset !== body.byteLength) throw invalidXfExt(record, "trailing-property-data");
}

function applyXfExtension(
  style: XfStyle,
  extensionType: number,
  data: Uint8Array,
  palette: string[],
  record: BiffRecord,
): void {
  switch (extensionType) {
    case 0x0004:
      style.fill = {
        ...style.fill,
        foregroundColor: parseFullColor(data, palette, record),
      };
      return;
    case 0x0005:
      style.fill = {
        ...style.fill,
        backgroundColor: parseFullColor(data, palette, record),
      };
      return;
    case 0x0006:
      style.fill = parseXfExtGradient(data, palette, record);
      return;
    case 0x0007:
      style.border = setBorderColor(style.border, "top", parseFullColor(data, palette, record));
      return;
    case 0x0008:
      style.border = setBorderColor(style.border, "bottom", parseFullColor(data, palette, record));
      return;
    case 0x0009:
      style.border = setBorderColor(style.border, "left", parseFullColor(data, palette, record));
      return;
    case 0x000a:
      style.border = setBorderColor(style.border, "right", parseFullColor(data, palette, record));
      return;
    case 0x000b:
      style.border = setBorderColor(
        style.border,
        "diagonal",
        parseFullColor(data, palette, record),
      );
      return;
    case 0x000d:
      style.font = { ...style.font, color: parseFullColor(data, palette, record) };
      return;
    case 0x000e:
      style.font = { ...style.font, scheme: parseFontScheme(data, record) };
      return;
    case 0x000f:
      style.alignment = {
        ...style.alignment,
        indent: parseIndent(data, record),
      };
      return;
    default:
      throw invalidXfExt(record, `unsupported-property-${extensionType}`);
  }
}

function setBorderColor(
  border: BorderSideOptions | undefined,
  side: "top" | "bottom" | "left" | "right" | "diagonal",
  color: ColorOptions,
): BorderSideOptions {
  return {
    ...border,
    [side]: { ...border?.[side], color },
  };
}

function parseFullColor(body: Uint8Array, palette: string[], record: BiffRecord): ColorOptions {
  if (body.byteLength < 16) throw invalidXfExt(record, "truncated-full-color");
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const type = view.getUint16(0, true);
  const tintShade = view.getInt16(2, true);
  if (tintShade === -32768) throw invalidXfExt(record, "invalid-color-tint");
  const value = view.getUint32(4, true);
  const color: ColorOptions =
    type === 0
      ? { auto: true }
      : type === 1
        ? { indexed: value }
        : type === 2
          ? { rgb: rgbaToArgb(value) }
          : type === 3
            ? { theme: value }
            : type === 4
              ? {}
              : (() => {
                  throw invalidXfExt(record, `unsupported-color-type-${type}`);
                })();
  if (tintShade !== 0 && type !== 0 && type !== 4) color.tint = tintShade / 32767;
  if (type === 1 && color.indexed !== undefined && color.indexed >= 8) {
    const customColor = palette[color.indexed - 8];
    if (customColor !== undefined) color.rgb = customColor as `FF${string}`;
  }
  return color;
}

function parseFontScheme(body: Uint8Array, record: BiffRecord): FontOptions["scheme"] {
  if (body.byteLength < 1) throw invalidXfExt(record, "truncated-font-scheme");
  const value =
    body.byteLength >= 2
      ? new DataView(body.buffer, body.byteOffset, body.byteLength).getUint16(0, true)
      : body[0]!;
  if (value === 1) return "major";
  if (value === 2) return "minor";
  if (value === 0 || value === 0xff) return "none";
  throw invalidXfExt(record, `unsupported-font-scheme-${value}`);
}

function parseIndent(body: Uint8Array, record: BiffRecord): number {
  if (body.byteLength < 2) throw invalidXfExt(record, "truncated-indent");
  const indent = new DataView(body.buffer, body.byteOffset, body.byteLength).getUint16(0, true);
  if (indent > 250) throw invalidXfExt(record, `invalid-indent-${indent}`);
  return indent;
}

function parseXfExtGradient(
  body: Uint8Array,
  palette: string[],
  record: BiffRecord,
): CellFillOptions {
  if (body.byteLength < 48) throw invalidXfExt(record, "truncated-gradient");
  const view = new DataView(body.buffer, body.byteOffset, body.byteLength);
  const gradientType = view.getUint32(0, true);
  if (gradientType > 1) throw invalidXfExt(record, `unsupported-gradient-type-${gradientType}`);
  const stopCount = view.getUint32(44, true);
  if (stopCount > 256 || 48 + stopCount * 22 > body.byteLength)
    throw invalidXfExt(record, "invalid-gradient-stops");
  const stops: NonNullable<CellFillOptions["stops"]> = [];
  let offset = 48;
  for (let index = 0; index < stopCount; index++) {
    const type = view.getUint16(offset, true);
    const value = view.getUint32(offset + 2, true);
    const position = view.getFloat64(offset + 6, true);
    const tint = view.getFloat64(offset + 14, true);
    if (position < 0 || position > 1 || tint < -1 || tint > 1)
      throw invalidXfExt(record, "invalid-gradient-stop");
    const color = parseCompactColor(type, value, palette, tint, record);
    stops.push({ position, color });
    offset += 22;
  }
  return {
    type: "gradient",
    gradientType: gradientType === 0 ? "linear" : "path",
    gradientDegree: view.getFloat64(4, true),
    gradientLeft: view.getFloat64(12, true),
    gradientRight: view.getFloat64(20, true),
    gradientTop: view.getFloat64(28, true),
    gradientBottom: view.getFloat64(36, true),
    stops,
  };
}

function parseCompactColor(
  type: number,
  value: number,
  palette: string[],
  tint: number,
  record: BiffRecord,
): ColorOptions {
  const color: ColorOptions =
    type === 0
      ? { auto: true }
      : type === 1
        ? { indexed: value }
        : type === 2
          ? { rgb: rgbaToArgb(value) }
          : type === 3
            ? { theme: value }
            : type === 4
              ? {}
              : (() => {
                  throw invalidXfExt(record, `unsupported-color-type-${type}`);
                })();
  if (type === 1 && color.indexed !== undefined && color.indexed >= 8) {
    const customColor = palette[color.indexed - 8];
    if (customColor !== undefined) color.rgb = customColor as `FF${string}`;
  }
  if (tint !== 0 && type !== 0 && type !== 4) color.tint = tint;
  return color;
}

function rgbaToArgb(value: number): `FF${string}` {
  const channel = (part: number): string => part.toString(16).padStart(2, "0");
  return `${channel((value >>> 24) & 0xff)}${channel(value & 0xff)}${channel(
    (value >>> 8) & 0xff,
  )}${channel((value >>> 16) & 0xff)}`.toUpperCase() as `FF${string}`;
}

function invalidXfExt(record: BiffRecord, reason: string): XlsParseError {
  return new XlsParseError(`Invalid legacy XLS file: ${reason}`, {
    part: "stream" as const,
    path: "Workbook/globals/XFExt",
    recordType: record.code,
    recordName: "XFExt",
    offset: record.start,
    length: record.body.byteLength,
    reason,
  });
}

function parseUsedAttributes(value: number): XfStyle["usedAttributes"] {
  return {
    font: (value & 0x0008) !== 0,
    format: (value & 0x0004) !== 0,
    alignment: (value & 0x0010) !== 0,
    border: (value & 0x0020) !== 0,
    fill: (value & 0x0040) !== 0,
    protection: (value & 0x0080) !== 0,
  };
}

function parseAlignment(
  body: Uint8Array,
  view: DataView,
  version: BiffVersion,
): AlignmentOptions | undefined {
  const alignmentOptions = view.getUint16(6, true);
  const horizontal = HORIZONTAL_ALIGNMENT[alignmentOptions & 0x0007];
  const vertical = VERTICAL_ALIGNMENT[(alignmentOptions & 0x0070) >> 4];
  if (horizontal === undefined || vertical === undefined) throw invalidXfField("unknown alignment");

  const alignment: AlignmentOptions = {};
  if (horizontal !== "general") alignment.horizontal = horizontal;
  if (vertical !== "top") alignment.vertical = vertical;
  if ((alignmentOptions & 0x0008) !== 0) alignment.wrapText = true;
  if (version === 8) {
    const rotation = body[7] ?? 0;
    if (rotation !== 0) alignment.textRotation = rotation;
    const indentationOptions = view.getUint16(8, true);
    if ((indentationOptions & 0x000f) !== 0) alignment.indent = indentationOptions & 0x000f;
    if ((indentationOptions & 0x0010) !== 0) alignment.shrinkToFit = true;
    if ((indentationOptions & 0x0020) !== 0) alignment.mergeCell = true;
    if ((alignmentOptions & 0x0080) !== 0) alignment.justifyLastLine = true;
    const readingOrder = (indentationOptions & 0x00c0) >> 6;
    if (readingOrder > 2) throw invalidXfField(`unknown reading order ${readingOrder}`);
    if (readingOrder !== 0) alignment.readingOrder = readingOrder;
  } else {
    const orientation = (alignmentOptions >> 8) & 0x0003;
    if (orientation === 1) alignment.textRotation = 255;
    if (orientation === 2) alignment.textRotation = 90;
    if (orientation === 3) alignment.textRotation = 180;
  }
  return Object.keys(alignment).length > 0 ? alignment : undefined;
}

function parseBorder(
  view: DataView,
  version: BiffVersion,
  palette: string[],
): BorderSideOptions | undefined {
  const border: BorderSideOptions = {};
  const setSide = (
    side: "left" | "right" | "top" | "bottom" | "diagonal",
    styleValue: number,
    color?: ColorOptions,
  ) => {
    const style = BORDER_STYLES[styleValue];
    if (style === undefined) throw invalidXfField(`unknown border style ${styleValue}`);
    if (styleValue === 0) return;
    border[side] = { style, ...(color ? { color } : {}) };
  };

  if (version === 8) {
    const borderOptions = view.getUint32(10, true);
    const paletteOptions = view.getUint32(12, true);
    const additionalOptions = view.getUint32(14, true);
    setSide("left", borderOptions & 0x000f, parseColor(paletteOptions & 0x007f, palette));
    setSide(
      "right",
      (borderOptions & 0x00f0) >> 4,
      parseColor((paletteOptions & 0x3f80) >> 7, palette),
    );
    setSide("top", (borderOptions & 0x0f00) >> 8, parseColor(additionalOptions & 0x7f, palette));
    setSide(
      "bottom",
      (borderOptions & 0xf000) >> 12,
      parseColor((additionalOptions & 0x3f80) >> 7, palette),
    );
    const diagonalStyle = (additionalOptions & 0x01e00000) >> 21;
    if (diagonalStyle !== 0)
      setSide(
        "diagonal",
        diagonalStyle,
        parseColor((additionalOptions & 0x001fc000) >> 14, palette),
      );
    if ((borderOptions & 0x40000000) !== 0) border.diagonalUp = true;
    if ((borderOptions & 0x80000000) !== 0) border.diagonalDown = true;
  } else {
    const areaOptions = view.getUint32(8, true);
    const borderOptions = view.getUint32(12, true);
    setSide(
      "top",
      borderOptions & 0x00000007,
      parseColor((borderOptions & 0x0000fe00) >> 9, palette),
    );
    setSide(
      "left",
      (borderOptions & 0x00000038) >> 3,
      parseColor((borderOptions & 0x007f0000) >> 16, palette),
    );
    setSide(
      "right",
      (borderOptions & 0x000001c0) >> 6,
      parseColor((borderOptions & 0x3f800000) >> 23, palette),
    );
    setSide(
      "bottom",
      (areaOptions & 0x01c00000) >> 22,
      parseColor((areaOptions & 0xfe000000) >>> 25, palette),
    );
  }
  return Object.keys(border).length > 0 ? border : undefined;
}

function parseFill(
  view: DataView,
  version: BiffVersion,
  palette: string[],
): CellFillOptions | undefined {
  const areaOptions = version === 8 ? view.getUint16(18, true) : view.getUint32(8, true);
  const foregroundColor = parseColor(areaOptions & 0x007f, palette);
  const backgroundColor = parseColor((areaOptions & 0x3f80) >> 7, palette);
  const patternValue =
    version === 8
      ? (view.getUint32(14, true) & 0xfc000000) >>> 26
      : (areaOptions & 0x003f0000) >> 16;
  const patternType = PATTERN_TYPES[patternValue];
  if (patternType === undefined) throw invalidXfField(`unknown fill pattern ${patternValue}`);
  if (patternType === "none") return undefined;
  return { type: "pattern", patternType, foregroundColor, backgroundColor };
}

function parseColor(index: number, palette: string[]): ColorOptions | undefined {
  if (index >= 8 && index - 8 < palette.length) return { rgb: palette[index - 8] as `FF${string}` };
  if (index > 0x4e) throw invalidXfField(`unknown color index ${index}`);
  return { indexed: index };
}

function invalidXfField(reason: string): XlsParseError {
  return new XlsParseError(`Invalid legacy XLS file: ${reason}`, {
    part: "stream",
    path: "Workbook/globals/XF",
    recordType: 0x00e0,
    recordName: "XF",
    reason,
  });
}
