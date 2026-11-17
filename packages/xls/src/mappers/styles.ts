import type {
  AlignmentOptions,
  BorderOptions,
  BorderSideOptions,
  CellFillOptions,
  ColorOptions,
  FontOptions,
} from "@office-open/xlsx";

import { XlsParseError } from "../errors";
import type { BiffRecord, BiffVersion } from "../records";
import { decodeBiffString, readBiff5String, readBiff8String } from "./biff";
import type { XfStyle } from "./models";
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
