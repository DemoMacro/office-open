import {
  attributeNumber,
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  escapeText,
  lengthToEmu,
  textOf,
  xmlElement,
} from "@office-open/odf";
import type {
  CellOptions,
  ColumnOptions,
  CommentOptions,
  MergeCellOptions,
  RowOptions,
  StyleOptions,
  WorksheetOptions,
} from "@office-open/xlsx";
import type { Element } from "@office-open/xml";

import { cellDecorationXml } from "./comments";
import { cellGraphicXml, type OdsChartFrame } from "./drawing-chart";
import { OdsParseError } from "./error";
import { odfDateValue } from "./numbering";
import { addCellStyle, addDimensionStyle, type DimensionStyle } from "./styles";
import { assertElementNames, unknownOdsElement } from "./workbook";

export const COLUMN_CONTAINERS = new Set([
  "table:table-columns",
  "table:table-header-columns",
  "table:table-column-group",
]);

export const ROW_CONTAINERS = new Set([
  "table:table-rows",
  "table:table-header-rows",
  "table:table-row-group",
]);

const MAX_COLUMNS = 16384;
const MAX_ROWS = 1048576;
const MAX_MATERIALIZED_EMPTY_REPEATS = 100;

function repeatCount(
  element: Element,
  name: string,
  maximum: number = Number.MAX_SAFE_INTEGER,
): number {
  const raw = attributeNumber(element, name);
  if (raw === undefined) return 1;
  if (!Number.isInteger(raw) || raw < 1 || raw > maximum) {
    throw new OdsParseError(
      "Invalid OpenDocument spreadsheet repeat count",
      "content.xml",
      "office:body/office:spreadsheet/table:table",
      name,
      "invalid-repeat-count",
    );
  }
  return raw;
}

function hasCellContent(cell: CellOptions): boolean {
  return Object.keys(cell).some((key) => key !== "reference" && key !== "style");
}

function hasVisibleCellStyle(style: StyleOptions | undefined): boolean {
  const border = style?.border;
  return (
    style?.fill !== undefined ||
    [border?.top, border?.bottom, border?.left, border?.right].some(
      (side) => side?.style !== undefined && side.style !== "none",
    )
  );
}

function isLargeEmptyRepeat(repeat: number, hasData: boolean): boolean {
  return !hasData && repeat > MAX_MATERIALIZED_EMPTY_REPEATS;
}

export function worksheetXml(
  worksheet: WorksheetOptions,
  index: number,
  styles: string[],
  embeddedCharts: OdsChartFrame[],
): string {
  const merges = (worksheet.mergeCells ?? []).flatMap((merge) => {
    const match = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/u.exec(merge.ref);
    if (!match) return [];
    const firstColumn = columnNumber(match[1]!);
    const firstRow = Number(match[2]!);
    const lastColumn = columnNumber(match[3]!);
    const lastRow = Number(match[4]!);
    if (firstColumn > lastColumn || firstRow > lastRow) return [];
    return [{ firstColumn, firstRow, lastColumn, lastRow }];
  });
  const explicitColumns = (worksheet.columns ?? []).flatMap((column) => {
    const count = Math.max(1, (column.max ?? column.min) - column.min + 1);
    const styleName =
      column.width === undefined && !column.hidden
        ? undefined
        : addDimensionStyle({ width: column.width, hidden: column.hidden }, styles, "column");
    return Array.from({ length: count }, () =>
      xmlElement("table:table-column", { "table:style-name": styleName }),
    );
  });
  let maxRowCells = 0;
  for (const row of worksheet.rows ?? []) {
    maxRowCells = Math.max(maxRowCells, row.cells?.length ?? 0);
  }
  const columnCount = Math.max(explicitColumns.length, maxRowCells, 1);
  const columns = concatenate(
    Array.from(
      { length: columnCount },
      (_, column) => explicitColumns[column] ?? xmlElement("table:table-column"),
    ),
  );
  const bodyRows = (worksheet.rows ?? []).map((row, rowIndex) => {
    const styleName =
      row.height === undefined && !row.hidden
        ? undefined
        : addDimensionStyle(
            { height: typeof row.height === "number" ? row.height : undefined, hidden: row.hidden },
            styles,
            "row",
          );
    const rowNumber = row.rowNumber ?? rowIndex + 1;
    const cells = (row.cells ?? []).map((cell, cellIndex) => ({
      cell,
      column: cell.reference ? columnNumber(cell.reference.replace(/\d+$/u, "")) : cellIndex + 1,
    }));
    for (const merge of merges) {
      if (rowNumber < merge.firstRow || rowNumber > merge.lastRow) continue;
      for (let column = merge.firstColumn; column <= merge.lastColumn; column += 1) {
        const reference = `${cellAddress(column, rowNumber)}`;
        if (cells.some((entry) => entry.cell.reference === reference)) continue;
        cells.push({ cell: { reference, covered: { reference } }, column });
      }
    }
    cells.sort((left, right) => left.column - right.column);
    return xmlElement(
      "table:table-row",
      { "table:style-name": styleName },
      cells.map(({ cell }) => {
        const start = merges.find(
          (merge) =>
            merge.firstRow === rowNumber &&
            merge.firstColumn ===
              (cell.reference ? columnNumber(cell.reference.replace(/\d+$/u, "")) : 0),
        );
        return cellXml(
          cell,
          styles,
          cellDecorationXml(worksheet, cell.reference ?? ""),
          start
            ? {
                columnsSpanned: start.lastColumn - start.firstColumn + 1,
                rowsSpanned: start.lastRow - start.firstRow + 1,
              }
            : undefined,
        );
      }),
    );
  });
  const rows = concatenate(
    bodyRows.length
      ? bodyRows
      : [xmlElement("table:table-row", undefined, [xmlElement("table:table-cell")])],
  );
  const sheetName = worksheet.name ?? `Sheet${index}`;
  const frames = embeddedCharts
    .filter((entry) => entry.worksheet === sheetName)
    .map((entry) =>
      xmlElement(
        "draw:frame",
        {
          "draw:name": entry.name,
          "svg:x": emuToLength(
            entry.chart.absoluteX ?? entry.chart.colOffset ?? entry.chart.col ?? 0,
          ),
          "svg:y": emuToLength(
            entry.chart.absoluteY ?? entry.chart.rowOffset ?? entry.chart.row ?? 0,
          ),
          "svg:width": emuToLength(entry.chart.extentCx ?? 400000),
          "svg:height": emuToLength(entry.chart.extentCy ?? 300000),
          ...(entry.chart.toCol !== undefined && entry.chart.toRow !== undefined
            ? {
                "table:end-cell-address": cellAddress(entry.chart.toCol, entry.chart.toRow),
                "table:end-x": emuToLength(entry.chart.toColOffset ?? 0),
                "table:end-y": emuToLength(entry.chart.toRowOffset ?? 0),
              }
            : {}),
        },
        [
          xmlElement("draw:object", {
            "xlink:href": entry.name,
            "xlink:type": "simple",
            "xlink:show": "embed",
            "xlink:actuate": "onLoad",
          }),
        ],
      ),
    );
  return xmlElement("table:table", { "table:name": sheetName }, [
    ...(frames.length ? [xmlElement("table:shapes", undefined, frames)] : []),
    columns,
    rows,
  ]);
}

function concatenate(parts: readonly string[]): string {
  let result = "";
  for (const part of parts) result += part;
  return result;
}

export function cellAddress(column: number, row: number): string {
  let value = column;
  let letters = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    value = Math.floor((value - 1) / 26);
  }
  return `${letters}${row}`;
}

export function columnNumber(column: string): number {
  let total = 0;
  for (let index = 0; index < column.length; index += 1) {
    total = total * 26 + column.charCodeAt(index) - 64;
  }
  return total;
}

export function cellXml(
  cell: CellOptions,
  styles: string[],
  decorations: string[] = [],
  span?: { columnsSpanned: number; rowsSpanned: number },
): string {
  if (cell.covered) {
    return xmlElement(
      "table:covered-table-cell",
      {
        "table:style-name":
          typeof cell.style === "object" ? addCellStyle(cell.style, styles) : undefined,
        "office:value-type": cell.covered.valueType,
      },
      cell.covered.text === undefined
        ? []
        : [xmlElement("text:p", undefined, [escapeText(cell.covered.text)])],
    );
  }
  const formula = typeof cell.formula === "string" ? cell.formula : cell.formula?.formula;
  const cached = cacheAttributes(cell.value);
  const styleOptions = typeof cell.style === "object" ? cell.style : undefined;
  const valueKind = numFmtKind(styleOptions?.numFmt);
  const type = valueKind === "time" ? "time" : valueKind === "currency" ? "currency" : cached.type;
  return xmlElement(
    "table:table-cell",
    {
      "table:style-name": styleOptions ? addCellStyle(styleOptions, styles) : undefined,
      "table:number-columns-spanned": span?.columnsSpanned,
      "table:number-rows-spanned": span?.rowsSpanned,
      "table:formula": formula ? `of:=${formula}` : undefined,
      "office:value-type": type,
      "office:value": cached.type === "float" ? cached.value : undefined,
      "office:currency":
        valueKind === "currency"
          ? /^\[\$([^-\]]+)/u.exec(styleOptions?.numFmt ?? "")?.[1]
          : undefined,
      "office:time-value":
        type === "time" ? serialToDuration(Number(cached.value ?? 0)) : undefined,
      "office:boolean-value": cached.type === "boolean" ? cached.value : undefined,
      "office:date-value": cached.type === "date" ? cached.value : undefined,
    },
    [
      ...(cached.type === "string"
        ? [xmlElement("text:p", undefined, [escapeText(String(cached.value ?? ""))])]
        : []),
      ...decorations,
      ...(cell.graphics ?? []).map((graphic) => cellGraphicXml(graphic)),
    ],
  );
}

export function cacheAttributes(value: CellOptions["value"]): {
  type?: string;
  value?: string | number | boolean;
} {
  if (typeof value === "number") return { type: "float", value };
  if (typeof value === "boolean") return { type: "boolean", value };
  if (value instanceof Date) return { type: "date", value: odfDateValue(value) };
  if (typeof value === "string") return { type: "string", value };
  return {};
}

function styleWithNumFmt(style: CellOptions["style"], numFmt: string): CellOptions["style"] {
  if (typeof style === "number") return style;
  return { ...style, numFmt: style?.numFmt ?? numFmt };
}

function durationToSerial(value: string | undefined): number {
  if (value === undefined) return 0;
  const match = /^(-?)P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:([\d.]+)S)?)?$/u.exec(value);
  if (!match) return 0;
  const sign = match[1] === "-" ? -1 : 1;
  const days = Number(match[2] ?? 0);
  const hours = Number(match[3] ?? 0);
  const minutes = Number(match[4] ?? 0);
  const seconds = Number(match[5] ?? 0);
  return sign * (days + hours / 24 + minutes / 1440 + seconds / 86400);
}

function serialToDuration(value: number): string {
  const negative = value < 0;
  let remaining = Math.abs(value) * 86400;
  const hours = Math.floor(remaining / 3600);
  remaining -= hours * 3600;
  const minutes = Math.floor(remaining / 60);
  const seconds = remaining - minutes * 60;
  const secondText = Number.isInteger(seconds) ? String(seconds) : seconds.toFixed(6);
  return `${negative ? "-" : ""}PT${hours}H${minutes}M${secondText}S`;
}

function numFmtKind(numFmt: string | undefined): "currency" | "time" | undefined {
  if (numFmt === undefined) return undefined;
  if (/^\[\$([^-\]]+)(?:-[^\]]+)?\]/u.test(numFmt)) return "currency";
  if (/\[?h+\]?[^ymd]*m+[^ymd]*(?:s|$)/iu.test(numFmt)) return "time";
  return undefined;
}

export function worksheet(
  table: Element,
  index: number,
  dimensions: Map<string, DimensionStyle>,
  cellStyles: Map<string, StyleOptions>,
): WorksheetOptions {
  const comments: CommentOptions[] = [];
  const mergeCells: MergeCellOptions[] = [];
  validateTableChildren(table);
  const columns: ColumnOptions[] = [];
  let columnNumber = 1;
  for (const column of tableColumns(table)) {
    const repeat = repeatCount(column, "table:number-columns-repeated", MAX_COLUMNS);
    const parsed = parseColumn(column, columnNumber, dimensions);
    parsed.max = columnNumber + repeat - 1;
    const previous = columns[columns.length - 1];
    if (
      previous &&
      previous.max === columnNumber - 1 &&
      previous.width === parsed.width &&
      previous.hidden === parsed.hidden
    ) {
      previous.max = parsed.max;
    } else {
      columns.push(parsed);
    }
    columnNumber += repeat;
  }
  const parsedRows = parseRows(table, dimensions, cellStyles, comments, mergeCells);
  const worksheetOptions: WorksheetOptions = {
    name: attributeString(table, "table:name") ?? `Sheet${index}`,
    ...(columns.length ? { columns } : {}),
    ...(parsedRows.defaultRowHeight === undefined
      ? {}
      : { sheetFormat: { defaultRowHeight: parsedRows.defaultRowHeight } }),
    rows: parsedRows.rows,
  };
  return {
    ...worksheetOptions,
    ...(comments.length ? { comments } : {}),
    ...(mergeCells.length ? { mergeCells } : {}),
  };
}

export function validateTableChildren(table: Element): void {
  for (const child of table.elements ?? []) {
    if (child.type !== "element" || !child.name) continue;
    if (
      child.name === "table:table-column" ||
      child.name === "table:table-row" ||
      child.name === "office:forms" ||
      child.name === "table:shapes" ||
      child.name === "draw:frame" ||
      COLUMN_CONTAINERS.has(child.name) ||
      ROW_CONTAINERS.has(child.name)
    ) {
      continue;
    }
    throw unknownOdsElement(
      child,
      `/office:document-content/office:body/office:spreadsheet/table:table[@table:name="${attributeString(table, "table:name") ?? ""}"]`,
      child.name,
      "element has no canonical WorksheetOptions mapping",
    );
  }
}

export function tableColumns(table: Element): Element[] {
  return (table.elements ?? []).flatMap((child) => {
    if (child.name === "table:table-column") return [child];
    if (COLUMN_CONTAINERS.has(child.name ?? "")) return tableColumns(child);
    return [];
  });
}

export function tableRows(table: Element): Element[] {
  return (table.elements ?? []).flatMap((child) => {
    if (child.name === "table:table-row") return [child];
    if (ROW_CONTAINERS.has(child.name ?? "")) return tableRows(child);
    return [];
  });
}

export function parseColumn(
  element: Element,
  column: number,
  dimensions: Map<string, DimensionStyle>,
): ColumnOptions {
  const style = dimensions.get(attributeString(element, "table:style-name") ?? "");
  return { min: column, max: column, width: style?.width, hidden: style?.hidden };
}

function parseRows(
  table: Element,
  dimensions: Map<string, DimensionStyle>,
  cellStyles: Map<string, StyleOptions>,
  comments: CommentOptions[],
  mergeCells: MergeCellOptions[],
): { rows: RowOptions[]; defaultRowHeight?: number } {
  const rows: RowOptions[] = [];
  let defaultRowHeight: number | undefined;
  let rowNumber = 1;
  for (const row of tableRows(table)) {
    const repeat = repeatCount(row, "table:number-rows-repeated");
    const parsed = parseRow(row, rowNumber, dimensions, cellStyles, comments, mergeCells);
    const hasRowData = (parsed.cells ?? []).some(hasCellContent) || parsed.hidden === true;
    if (repeat > MAX_ROWS && hasRowData) {
      throw new OdsParseError(
        "Invalid OpenDocument spreadsheet repeat count",
        "content.xml",
        "office:body/office:spreadsheet/table:table",
        "table:number-rows-repeated",
        "invalid-repeat-count",
      );
    }
    if (isLargeEmptyRepeat(repeat, hasRowData)) {
      if (typeof parsed.height === "number" && defaultRowHeight === undefined) {
        defaultRowHeight = parsed.height;
      }
      rowNumber += repeat;
      continue;
    }
    if (repeat > 1) {
      for (let index = 0; index < repeat; index += 1) {
        const currentRow = rowNumber + index;
        rows.push({
          ...parsed,
          rowNumber: currentRow,
          cells: (parsed.cells ?? []).map((cell) => renumberCell(cell, currentRow)),
        });
      }
      rowNumber += repeat;
      continue;
    }
    rows.push(parsed);
    rowNumber += repeat;
  }
  return { rows, defaultRowHeight };
}

function renumberCell(cell: CellOptions, rowNumber: number): CellOptions {
  const next = {
    ...cell,
    reference: replaceRowReference(cell.reference ?? "", rowNumber),
    ...(cell.covered
      ? {
          covered: {
            ...cell.covered,
            reference: replaceRowReference(cell.covered.reference, rowNumber),
          },
        }
      : {}),
  };
  return next;
}

function replaceRowReference(reference: string, rowNumber: number): string {
  const match = /^([A-Z]+)(\d+)$/u.exec(reference);
  return match ? `${match[1]}${rowNumber}` : reference;
}

export function parseRow(
  row: Element,
  rowNumber: number,
  dimensions: Map<string, DimensionStyle>,
  cellStyles: Map<string, StyleOptions>,
  comments: CommentOptions[],
  mergeCells: MergeCellOptions[] = [],
): RowOptions {
  const style = dimensions.get(attributeString(row, "table:style-name") ?? "");
  const cellStartColumns: number[] = [];
  let columnNumber = 1;
  for (const cell of row.elements ?? []) {
    if (cell.type !== "element" || !cell.name) continue;
    const repeat = repeatCount(cell, "table:number-columns-repeated");
    cellStartColumns.push(columnNumber);
    columnNumber += repeat;
  }
  return {
    rowNumber,
    height: style?.height,
    hidden: style?.hidden,
    cells: (row.elements ?? [])
      .flatMap((cell, cellIndex) => {
        const repeat = repeatCount(cell, "table:number-columns-repeated");
        const startColumn = cellStartColumns[cellIndex]!;
        if (cell.type !== "element" || !cell.name) return undefined;
        if (cell.name === "table:covered-table-cell") {
          assertElementNames(
            cell,
            ["table:style-name", "office:value-type", "calcext:value-type"],
            ["text:p"],
          );
          const coveredCell = {
            reference: `${columnName(startColumn)}${rowNumber}`,
            covered: {
              reference: `${columnName(startColumn)}${rowNumber}`,
              ...(childNamed(cell, "text:p") ? { text: textOf(childNamed(cell, "text:p")) } : {}),
              ...(attributeString(cell, "office:value-type")
                ? {
                    valueType: attributeString(cell, "office:value-type") as
                      | "float"
                      | "string"
                      | "date"
                      | "boolean"
                      | "percentage"
                      | "currency"
                      | "time",
                  }
                : {}),
            },
            ...(cellStyles.get(attributeString(cell, "table:style-name") ?? "")
              ? { style: cellStyles.get(attributeString(cell, "table:style-name") ?? "") }
              : {}),
          };
          if (repeat > MAX_COLUMNS) {
            if (
              !childNamed(cell, "text:p") &&
              !attributeString(cell, "office:value-type") &&
              !hasVisibleCellStyle(coveredCell.style)
            ) {
              return [];
            }
            throw new OdsParseError(
              "Invalid OpenDocument spreadsheet repeat count",
              "content.xml",
              "office:body/office:spreadsheet/table:table",
              "table:number-columns-repeated",
              "invalid-repeat-count",
            );
          }
          return Array.from({ length: repeat }, (_, repeatedIndex) => ({
            ...coveredCell,
            reference: `${columnName(startColumn + repeatedIndex)}${rowNumber}`,
            covered: {
              ...coveredCell.covered,
              reference: `${columnName(startColumn + repeatedIndex)}${rowNumber}`,
            },
          }));
        }
        if (cell.name !== "table:table-cell") {
          throw unknownOdsElement(
            cell,
            `/office:document-content/office:body/office:spreadsheet/table:table/table:table-row[${rowNumber}]`,
            cell.name,
            "element has no canonical CellOptions mapping",
          );
        }
        const firstCell = parseCell(cell, rowNumber, startColumn, cellStyles, []);
        const hasMergeSpan =
          (attributeNumber(cell, "table:number-columns-spanned") ?? 1) > 1 ||
          (attributeNumber(cell, "table:number-rows-spanned") ?? 1) > 1;
        const hasCellSurface =
          hasCellContent(firstCell) ||
          hasMergeSpan ||
          hasVisibleCellStyle(typeof firstCell.style === "number" ? undefined : firstCell.style);
        if (repeat > MAX_COLUMNS && hasCellSurface) {
          throw new OdsParseError(
            "Invalid OpenDocument spreadsheet repeat count",
            "content.xml",
            "office:body/office:spreadsheet/table:table",
            "table:number-columns-repeated",
            "invalid-repeat-count",
          );
        }
        if (isLargeEmptyRepeat(repeat, hasCellSurface)) return [];
        return Array.from({ length: repeat }, (_, repeatedIndex) => {
          const column = startColumn + repeatedIndex;
          const parsed = parseCell(cell, rowNumber, column, cellStyles, mergeCells);
          for (const child of cell.elements ?? []) {
            if (child.name === "office:annotation") {
              comments.push({
                cell: parsed.reference ?? "",
                author: attributeString(child, "dc:creator") ?? "",
                text: childrenNamed(child, "text:p")
                  .map((paragraph) => textOf(paragraph))
                  .join("\n"),
              });
            }
            if (child.name === "draw:frame") {
              const object =
                childNamed(child, "draw:object") ?? childNamed(child, "draw:object-ole");
              const image = childNamed(child, "draw:image");
              const href = attributeString(object ?? image, "xlink:href");
              const description = textOf(childNamed(child, "svg:desc"));
              if (!href) {
                throw unknownOdsElement(
                  child,
                  `/office:document-content/office:body/office:spreadsheet/table:table/table:table-row[${rowNumber}]/table:table-cell[${column}]`,
                  child.name!,
                  "graphic href is required",
                );
              }
              assertElementNames(
                child,
                [
                  "draw:name",
                  "svg:x",
                  "svg:y",
                  "svg:width",
                  "svg:height",
                  "draw:z-index",
                  "table:end-cell-address",
                  "draw:id",
                  "table:end-x",
                  "table:end-y",
                  "draw:style-name",
                  "draw:text-style-name",
                  "xmlns:presentation",
                  "style:rel-width",
                  "style:rel-height",
                ],
                ["draw:image", "draw:object", "draw:object-ole", "svg:desc"],
              );
              parsed.graphics = [
                ...(parsed.graphics ?? []),
                object
                  ? {
                      type: "object",
                      href: href.replace(/^\.\//, "").replace(/^\//, ""),
                      ...(attributeString(child, "draw:name")
                        ? { name: attributeString(child, "draw:name") }
                        : {}),
                      ...(description ? { description } : {}),
                      x: lengthToEmu(attributeString(child, "svg:x")),
                      y: lengthToEmu(attributeString(child, "svg:y")),
                      width: lengthToEmu(attributeString(child, "svg:width")),
                      height: lengthToEmu(attributeString(child, "svg:height")),
                      ...(attributeNumber(child, "draw:z-index") !== undefined
                        ? { zIndex: attributeNumber(child, "draw:z-index") }
                        : {}),
                      ...(attributeString(child, "table:end-cell-address")
                        ? { endCellAddress: attributeString(child, "table:end-cell-address") }
                        : {}),
                      ...(attributeString(child, "draw:id")
                        ? { id: attributeString(child, "draw:id") }
                        : {}),
                      ...(attributeString(child, "table:end-x") !== undefined
                        ? { endX: lengthToEmu(attributeString(child, "table:end-x")) }
                        : {}),
                      ...(attributeString(child, "table:end-y") !== undefined
                        ? { endY: lengthToEmu(attributeString(child, "table:end-y")) }
                        : {}),
                      ...(attributeString(child, "draw:style-name")
                        ? { styleName: attributeString(child, "draw:style-name") }
                        : {}),
                      ...(attributeString(child, "draw:text-style-name")
                        ? { textStyleName: attributeString(child, "draw:text-style-name") }
                        : {}),
                      ...(attributeString(child, "style:rel-width")
                        ? { relativeWidth: attributeString(child, "style:rel-width") }
                        : {}),
                      ...(attributeString(child, "style:rel-height")
                        ? { relativeHeight: attributeString(child, "style:rel-height") }
                        : {}),
                    }
                  : {
                      type: "image",
                      href: isExternalUrl(href) ? href : href.replace(/^\//, ""),
                      data: "",
                      ...(attributeString(child, "draw:name")
                        ? { name: attributeString(child, "draw:name") }
                        : {}),
                      ...(description ? { description } : {}),
                      x: lengthToEmu(attributeString(child, "svg:x")),
                      y: lengthToEmu(attributeString(child, "svg:y")),
                      width: lengthToEmu(attributeString(child, "svg:width")),
                      height: lengthToEmu(attributeString(child, "svg:height")),
                      ...(attributeNumber(child, "draw:z-index") !== undefined
                        ? { zIndex: attributeNumber(child, "draw:z-index") }
                        : {}),
                      ...(attributeString(child, "table:end-cell-address")
                        ? { endCellAddress: attributeString(child, "table:end-cell-address") }
                        : {}),
                      ...(attributeString(child, "draw:id")
                        ? { id: attributeString(child, "draw:id") }
                        : {}),
                      ...(attributeString(child, "table:end-x") !== undefined
                        ? { endX: lengthToEmu(attributeString(child, "table:end-x")) }
                        : {}),
                      ...(attributeString(child, "table:end-y") !== undefined
                        ? { endY: lengthToEmu(attributeString(child, "table:end-y")) }
                        : {}),
                      ...(attributeString(child, "draw:style-name")
                        ? { styleName: attributeString(child, "draw:style-name") }
                        : {}),
                      ...(attributeString(child, "draw:text-style-name")
                        ? { textStyleName: attributeString(child, "draw:text-style-name") }
                        : {}),
                      ...(attributeString(child, "style:rel-width")
                        ? { relativeWidth: attributeString(child, "style:rel-width") }
                        : {}),
                      ...(attributeString(child, "style:rel-height")
                        ? { relativeHeight: attributeString(child, "style:rel-height") }
                        : {}),
                    },
              ];
            }
          }
          return parsed;
        });
      })
      .filter((cell): cell is CellOptions => cell !== undefined),
  };
}

function isExternalUrl(value: string): boolean {
  return /^[a-z][a-z0-9+.-]*:/i.test(value);
}

export function parseCell(
  cell: Element,
  row: number,
  column: number,
  cellStyles: Map<string, StyleOptions>,
  mergeCells: MergeCellOptions[] = [],
): CellOptions {
  const result: CellOptions = { reference: `${columnName(column)}${row}` };
  const columnSpan = attributeNumber(cell, "table:number-columns-spanned") ?? 1;
  const rowSpan = attributeNumber(cell, "table:number-rows-spanned") ?? 1;
  if (
    !Number.isInteger(columnSpan) ||
    !Number.isInteger(rowSpan) ||
    columnSpan < 1 ||
    rowSpan < 1 ||
    column + columnSpan - 1 > MAX_COLUMNS ||
    row + rowSpan - 1 > MAX_ROWS
  ) {
    throw new OdsParseError(
      "Invalid OpenDocument spreadsheet cell span",
      "content.xml",
      `office:body/office:spreadsheet/table:table/table:table-row[${row}]`,
      cell.name ?? "",
      "invalid-cell-span",
    );
  }
  if (columnSpan > 1 || rowSpan > 1) {
    mergeCells.push({
      ref: `${columnName(column)}${row}:${columnName(column + columnSpan - 1)}${row + rowSpan - 1}`,
    });
  }
  for (const child of cell.elements ?? []) {
    if (
      child.type === "element" &&
      !["text:p", "office:annotation", "draw:frame"].includes(child.name ?? "")
    ) {
      throw unknownOdsElement(
        child,
        `/office:document-content/office:body/office:spreadsheet/table:table/table:table-row[${row}]/table:table-cell[${column}]`,
        child.name ?? "",
        "element has no canonical CellOptions mapping",
      );
    }
  }
  const style = cellStyles.get(attributeString(cell, "table:style-name") ?? "");
  if (style) result.style = style;
  const formula = attributeString(cell, "table:formula");
  if (formula) result.formula = formula.replace(/^of:=/, "");
  const valueType = attributeString(cell, "office:value-type");
  if (valueType === "float") result.value = attributeNumber(cell, "office:value");
  else if (valueType === "percentage") {
    result.value = attributeNumber(cell, "office:value");
    result.style = styleWithNumFmt(result.style, "0%");
  } else if (valueType === "currency") {
    result.value = attributeNumber(cell, "office:value");
    const currency = attributeString(cell, "office:currency");
    if (currency) result.style = styleWithNumFmt(result.style, `[$${currency}] #,##0.00`);
  } else if (valueType === "time") {
    result.value = durationToSerial(attributeString(cell, "office:time-value"));
    result.style = styleWithNumFmt(result.style, "[hh]:mm:ss");
  } else if (valueType === "boolean")
    result.value = attributeString(cell, "office:boolean-value") === "true";
  else if (valueType === "date")
    result.value = new Date(`${attributeString(cell, "office:date-value")}Z`);
  else if (valueType === "string") result.value = textOf(childNamed(cell, "text:p"));
  return result;
}

export function columnName(column: number): string {
  let result = "";
  let value = column;
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}
