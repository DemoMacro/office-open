import type { CellOptions, FontOptions, FormulaOptions, StyleOptions } from "@office-open/xlsx";

import { XlsParseError } from "../errors";
import { columnLabel, EXCEL_ERROR_LITERALS } from "../shared/cell-reference";
import type { SheetRow, SheetState, TxoTextPart, WorkbookState, XfStyle } from "./models";
const DEFAULT_FONT: FontOptions = {
  size: 10,
  bold: false,
  italic: false,
  color: { indexed: 0 },
  underline: false,
  font: "Arial",
};

export function decodeRk(value: number): number {
  const divideBy100 = (value & 0x01) !== 0;
  const isInteger = (value & 0x02) !== 0;
  let result: number;
  if (isInteger) {
    result = value >> 2;
  } else {
    const bytes = new Uint8Array(8);
    new DataView(bytes.buffer).setUint32(4, value & 0xfffffffc, true);
    result = new DataView(bytes.buffer).getFloat64(0, true);
  }
  return divideBy100 ? result / 100 : result;
}

export function concatBytes(parts: readonly Uint8Array[]): Uint8Array {
  const total = parts.reduce((length, part) => length + part.byteLength, 0);
  const bytes = new Uint8Array(total);
  let offset = 0;
  for (const part of parts) {
    bytes.set(part, offset);
    offset += part.byteLength;
  }
  return bytes;
}

export function decodeParts(parts: readonly TxoTextPart[]): string {
  return parts.map(({ encoding, data }) => new TextDecoder(encoding).decode(data)).join("");
}

export function reference(row: number, column: number): string {
  return `${columnLabel(column)}${row + 1}`;
}

export function ensureRow(sheet: SheetState, rowNumber: number): SheetRow {
  const existing = sheet.rows.get(rowNumber);
  if (existing) return existing;
  const row: SheetRow = { rowNumber: rowNumber + 1, cells: [] };
  sheet.rows.set(rowNumber, row);
  return row;
}

export function cellStyle(style: number, state: WorkbookState): number | StyleOptions {
  const xf = resolveXf(style, state, new Set());
  if (!xf) return style;
  const result: StyleOptions = {};
  const font = state.fonts[xf.fontIndex];
  if (font) {
    result.font = font;
  } else if (xf.fontIndex === state.fonts.length) {
    result.font = DEFAULT_FONT;
  } else {
    throw new XlsParseError("Invalid legacy XLS file: extended format references missing font", {
      part: "stream",
      path: "Workbook/globals/XF",
      recordType: 0x00e0,
      recordName: "XF",
      reason: "missing-font",
    });
  }
  const format = state.numberFormats.get(xf.numberFormatId);
  if (format) result.numFmt = format;
  if (xf.fill) result.fill = xf.fill;
  if (xf.border) result.border = xf.border;
  if (xf.alignment) result.alignment = xf.alignment;
  if (xf.protection) result.protection = xf.protection;
  if (xf.quotePrefix) result.quotePrefix = true;
  return result;
}

function resolveXf(
  styleIndex: number,
  state: WorkbookState,
  visited: Set<number>,
): XfStyle | undefined {
  const xf = state.styleTable.get(styleIndex);
  if (!xf || xf.isStyle || visited.has(styleIndex)) return xf;
  visited.add(styleIndex);
  const parent = resolveXf(xf.parentIndex, state, visited);
  if (!parent) return xf;
  const used = xf.usedAttributes;
  return {
    ...xf,
    fontIndex: used.font ? xf.fontIndex : parent.fontIndex,
    numberFormatId: used.format ? xf.numberFormatId : parent.numberFormatId,
    alignment: xf.alignment ?? parent.alignment,
    border: xf.border ?? parent.border,
    fill: xf.fill ?? parent.fill,
    protection: xf.protection ?? parent.protection,
  };
}

export function makeCell(
  row: number,
  column: number,
  style: number,
  value: CellOptions["value"],
  state?: WorkbookState,
): CellOptions {
  return {
    reference: reference(row, column),
    value,
    style: state ? cellStyle(style, state) : style,
  };
}

export function formulaOptions(
  expression?: string,
  overrides?: Partial<FormulaOptions>,
): FormulaOptions | undefined {
  return expression ? { formula: expression, ...overrides } : undefined;
}

export function errorLiteral(code: number): string {
  const literal = EXCEL_ERROR_LITERALS[code];
  if (!literal)
    throw new XlsParseError(
      `Invalid legacy XLS file: unknown formula error code 0x${code.toString(16)}`,
    );
  return literal;
}

function isStringResult(value: Uint8Array): boolean {
  return (
    value.byteLength === 8 &&
    value.subarray(0, 6).every((byte) => byte === 0) &&
    value[6] === 0xff &&
    value[7] === 0xff
  );
}

export function cachedFormulaResult(value: Uint8Array): CellOptions["value"] | undefined {
  if (isStringResult(value)) return undefined;
  const view = new DataView(value.buffer, value.byteOffset, value.byteLength);
  if (!value.subarray(0, 6).every((byte) => byte === 0)) return view.getFloat64(0, true);
  if (value[6] === 0x00) return view.getFloat64(0, true);
  if (value[6] === 0x01) return value[7] !== 0;
  if (value[6] === 0x02) return errorLiteral(value[7]!);
  if (value[6] === 0x03) return "";
  if (value[6] === 0xff) return undefined;
  return view.getFloat64(0, true);
}
