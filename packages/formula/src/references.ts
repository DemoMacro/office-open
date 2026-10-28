import type { FormulaCellCoordinate, FormulaDialect, FormulaRangeEdge } from "./ast";
import type { FormulaReferenceArea } from "./ast";

export const maximumRows = 1_048_576;
export const maximumColumns = 16_384;

export function columnToNumber(column: string): number {
  let value = 0;
  for (const character of column.toUpperCase()) {
    value = value * 26 + character.charCodeAt(0) - 64;
  }
  return value;
}

export function numberToColumn(column: number): string {
  if (column < 1 || column > maximumColumns)
    throw new RangeError(`column ${column} is outside A..XFD`);
  let value = column;
  let text = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    text = String.fromCharCode(65 + remainder) + text;
    value = Math.floor((value - 1) / 26);
  }
  return text;
}

export function formatCell(coordinate: FormulaCellCoordinate): string {
  return `${coordinate.absoluteColumn ? "$" : ""}${numberToColumn(coordinate.column)}${coordinate.absoluteRow ? "$" : ""}${coordinate.row}`;
}

export function formatRangeEdge(edge: FormulaRangeEdge): string {
  if (isWholeColumn(edge))
    return `${edge.absolute ? "$" : ""}${numberToColumn(edge.column)}:${edge.absolute ? "$" : ""}${numberToColumn(edge.column)}`;
  if (isWholeRow(edge))
    return `${edge.absolute ? "$" : ""}${edge.row}:${edge.absolute ? "$" : ""}${edge.row}`;
  return formatCell(edge);
}

export function formatReferenceArea(
  area: FormulaReferenceArea,
  dialect: FormulaDialect = "excel",
): string {
  const end = area.end ?? area.start;
  let range: string;
  if (isWholeColumn(area.start) && isWholeColumn(end)) {
    range = wholeColumnRange(area.start, end);
  } else if (isWholeRow(area.start) && isWholeRow(end)) {
    range = wholeRowRange(area.start, end);
  } else {
    if (
      isWholeColumn(area.start) ||
      isWholeColumn(end) ||
      isWholeRow(area.start) ||
      isWholeRow(end)
    )
      throw new RangeError("range edges must use the same reference axis");
    range =
      formatCell(area.start) === formatCell(end)
        ? formatCell(area.start)
        : `${formatCell(area.start)}:${formatCell(end)}`;
  }
  const workbook =
    dialect === "excel"
      ? formatExcelWorkbook(area.workbook)
      : formatOpenFormulaWorkbook(area.workbook);
  const sheets =
    area.fromSheet === undefined
      ? ""
      : area.toSheet === undefined || area.toSheet === area.fromSheet
        ? formatSheetName(area.fromSheet)
        : `${formatSheetName(area.fromSheet)}:${formatSheetName(area.toSheet)}`;
  return formatArea(area, dialect, range, workbook, sheets);
}

function formatExcelWorkbook(workbook: FormulaReferenceArea["workbook"]): string {
  const identity =
    workbook?.path !== undefined
      ? workbook.path
      : workbook?.index !== undefined
        ? String(workbook.index)
        : workbook?.name;
  if (identity === undefined) return "";
  return /^[\w.-]+$/u.test(identity) ? `[${identity}]` : `['${identity.replaceAll("'", "''")}']`;
}

function formatOpenFormulaWorkbook(workbook: FormulaReferenceArea["workbook"]): string {
  if (workbook?.index !== undefined) return `${workbook.index}#`;
  const identity = workbook?.path ?? workbook?.name;
  if (identity === undefined) return "";
  return `'${identity.replaceAll("'", "''")}'#`;
}

function formatArea(
  area: FormulaReferenceArea,
  dialect: FormulaDialect,
  range: string,
  workbook = "",
  sheets = "",
): string {
  const { fromSheet, toSheet } = area;
  if (dialect === "excel") return `${workbook}${sheets}${sheets ? "!" : ""}${range}`;
  const [start, end] = range.split(":");
  const workbookPrefix = workbook;
  if (toSheet !== undefined && fromSheet !== undefined && toSheet !== fromSheet) {
    return `[${workbookPrefix}${formatSheetName(fromSheet)}.${start}:${formatSheetName(toSheet)}.${end ?? start}]`;
  }
  if (fromSheet !== undefined) {
    return `[${workbookPrefix}${formatSheetName(fromSheet)}.${start}${end === undefined ? "" : `:${formatSheetName(fromSheet)}.${end}`}]`;
  }
  return `[${workbookPrefix}.${start}${end === undefined ? "" : `:.${end}`}]`;
}

function wholeColumnRange(
  start: Extract<FormulaRangeEdge, { kind: "whole-column" }>,
  end: Extract<FormulaRangeEdge, { kind: "whole-column" }>,
): string {
  return `${formatWholeColumn(start)}:${formatWholeColumn(end)}`;
}

function wholeRowRange(
  start: Extract<FormulaRangeEdge, { kind: "whole-row" }>,
  end: Extract<FormulaRangeEdge, { kind: "whole-row" }>,
): string {
  return `${formatWholeRow(start)}:${formatWholeRow(end)}`;
}

function formatWholeColumn(edge: Extract<FormulaRangeEdge, { kind: "whole-column" }>): string {
  return `${edge.absolute ? "$" : ""}${numberToColumn(edge.column)}`;
}

function formatWholeRow(edge: Extract<FormulaRangeEdge, { kind: "whole-row" }>): string {
  return `${edge.absolute ? "$" : ""}${edge.row}`;
}

function formatSheetName(name: string): string {
  return /[^A-Za-z0-9_]/u.test(name) ? `'${name.replaceAll("'", "''")}'` : name;
}

export function shiftCell(
  coordinate: FormulaCellCoordinate,
  rows: number,
  columns: number,
): FormulaCellCoordinate {
  if (coordinate.absoluteRow && coordinate.absoluteColumn) return coordinate;
  const row = coordinate.absoluteRow ? coordinate.row : coordinate.row + rows;
  const column = coordinate.absoluteColumn ? coordinate.column : coordinate.column + columns;
  if (row < 1 || row > maximumRows || column < 1 || column > maximumColumns)
    throw new RangeError("shifted reference is outside the worksheet");
  return {
    ...coordinate,
    row,
    column,
  };
}

export function shiftRangeEdge(
  edge: FormulaRangeEdge,
  rows: number,
  columns: number,
): FormulaRangeEdge {
  if (isWholeColumn(edge)) {
    if (edge.absolute) return edge;
    const column = edge.column + columns;
    if (column < 1 || column > maximumColumns)
      throw new RangeError("shifted reference is outside the worksheet");
    return { ...edge, column };
  }
  if (isWholeRow(edge)) {
    if (edge.absolute) return edge;
    const row = edge.row + rows;
    if (row < 1 || row > maximumRows)
      throw new RangeError("shifted reference is outside the worksheet");
    return { ...edge, row };
  }
  return shiftCell(edge, rows, columns);
}

export function isWholeColumn(
  edge: FormulaRangeEdge,
): edge is Extract<FormulaRangeEdge, { kind: "whole-column" }> {
  return typeof edge === "object" && "kind" in edge && edge.kind === "whole-column";
}

export function isWholeRow(
  edge: FormulaRangeEdge,
): edge is Extract<FormulaRangeEdge, { kind: "whole-row" }> {
  return typeof edge === "object" && "kind" in edge && edge.kind === "whole-row";
}
