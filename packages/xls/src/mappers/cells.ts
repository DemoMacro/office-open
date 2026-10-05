import type { CellOptions, FormulaOptions, RowOptions, StyleOptions } from "@office-open/xlsx";

import { LegacyExcelError } from "../errors";
import type { SheetState, TxoTextPart, WorkbookState } from "./models";
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

export function columnLabel(column: number): string {
  let value = column;
  let label = "";
  do {
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26) - 1;
  } while (value >= 0);
  return label;
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

export function ensureRow(sheet: SheetState, rowNumber: number): RowOptions {
  const existing = sheet.rows.get(rowNumber);
  if (existing) return existing;
  const row: RowOptions = { rowNumber: rowNumber + 1, cells: [] };
  sheet.rows.set(rowNumber, row);
  return row;
}

export function cellStyle(style: number, state: WorkbookState): number | StyleOptions {
  const xf = state.styleTable.get(style);
  if (!xf) return style;
  const font = state.fonts[xf.fontIndex ?? -1];
  const format =
    xf.numberFormatId === undefined ? undefined : state.numberFormats.get(xf.numberFormatId);
  return { font, numFmt: format };
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
  const literals: { [code: number]: string } = {
    0x00: "#NULL!",
    0x07: "#DIV/0!",
    0x0f: "#VALUE!",
    0x17: "#REF!",
    0x1d: "#NAME?",
    0x24: "#NUM!",
    0x2a: "#N/A",
    0x2b: "#GETTING_DATA",
  };
  const literal = literals[code];
  if (!literal)
    throw new LegacyExcelError(
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
