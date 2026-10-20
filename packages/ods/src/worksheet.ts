import {
  attributeNumber,
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  escapeText,
  isExternalOdfUrl,
  lengthToEmu,
  spreadsheetColumnName,
  spreadsheetColumnNumber,
  textOf,
  xmlElement,
} from "@office-open/odf";
import type { GraphicStyle } from "@office-open/odf";
import type {
  CellOptions,
  CellRangeSourceOptions,
  ColumnOptions,
  CommentOptions,
  DetectiveHighlightedRangeOptions,
  DetectiveOperationOptions,
  DxfOptions,
  MergeCellOptions,
  RowOptions,
  SelectionOptions,
  StyleOptions,
  TableDetectiveOptions,
  WorksheetOptions,
} from "@office-open/xlsx";
import type { Element } from "@office-open/xml";

import { cellDecorationXml } from "./comments";
import { cellGraphicXml, type OdsChartFrame } from "./drawing-chart";
import { OdsParseError } from "./error";
import { odfDateValue } from "./numbering";
import {
  cellShapeXml,
  conditionalFormatsXml,
  parseDrawingShape,
  tableProtectionXml,
} from "./spreadsheet";
import { addCellStyle, addDimensionStyle, addTableStyle, type DimensionStyle } from "./styles";
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

type DetectiveHighlightedRangeDirection = Extract<
  DetectiveHighlightedRangeOptions,
  { direction: unknown }
>["direction"];

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
  localNamedExpressions = "",
  validationNames = new Map<string, string>(),
  dxfs: DxfOptions[] = [],
  formsXml = "",
): string {
  const conditionalStyleNames = new Map<number, string>();
  for (const rule of (worksheet.conditionalFormats ?? []).flatMap((format) => format.rules)) {
    const dxfId = rule.dxfId;
    if (dxfId === undefined || conditionalStyleNames.has(dxfId)) continue;
    const dxf = dxfs[dxfId];
    if (!dxf) continue;
    conditionalStyleNames.set(
      dxfId,
      addCellStyle(
        {
          font: dxf.font,
          fill: dxf.fill,
          border: dxf.border,
          numFmt: typeof dxf.numFmt === "string" ? dxf.numFmt : dxf.numFmt?.formatCode,
          alignment: dxf.alignment,
          protection: dxf.protection,
        },
        styles,
      ),
    );
  }
  const tableSheetName = worksheet.name ?? `Sheet${index}`;
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
          [
            ...cellDecorationXml(worksheet, cell.reference ?? ""),
            ...cellShapeXml(worksheet.shapes, cell.reference ?? "", styles),
          ],
          validationNames.get(`${tableSheetName}:${cell.reference ?? ""}`),
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
  const tableStyleName =
    worksheet.sheetView?.showFormulas !== undefined ||
    worksheet.sheetView?.showRowColHeaders !== undefined ||
    worksheet.state === "hidden"
      ? addTableStyle(
          {
            hidden: worksheet.state === "hidden",
            showFormulas: worksheet.sheetView?.showFormulas,
            showColumnHeaders: worksheet.sheetView?.showRowColHeaders,
            showRowHeaders: worksheet.sheetView?.showRowColHeaders,
          },
          styles,
        )
      : undefined;
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
  const selectionXml = worksheet.selection?.length ? gnumericSelectionXml(worksheet.selection) : [];
  return xmlElement(
    "table:table",
    { "table:name": sheetName, "table:style-name": tableStyleName },
    [
      formsXml,
      ...(frames.length ? [xmlElement("table:shapes", undefined, frames)] : []),
      tableProtectionXml(worksheet),
      columns,
      rows,
      conditionalFormatsXml(worksheet.conditionalFormats ?? [], sheetName, conditionalStyleNames),
      localNamedExpressions,
      ...selectionXml,
    ],
  );
}

function concatenate(parts: readonly string[]): string {
  let result = "";
  for (const part of parts) result += part;
  return result;
}

export function cellAddress(column: number, row: number): string {
  return `${spreadsheetColumnName(column)}${row}`;
}

export function columnNumber(column: string): number {
  return spreadsheetColumnNumber(column);
}

export function cellXml(
  cell: CellOptions,
  styles: string[],
  decorations: string[] = [],
  contentValidationName?: string,
  span?: { columnsSpanned: number; rowsSpanned: number },
): string {
  if (cell.covered) {
    return xmlElement(
      "table:covered-table-cell",
      {
        "table:style-name":
          typeof cell.style === "object" ? addCellStyle(cell.style, styles) : undefined,
        "office:value-type": cell.covered.valueType,
        "table:content-validation-name": contentValidationName,
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
  let type = cached.type;
  if (valueKind === "time") type = "time";
  else if (valueKind === "currency") type = "currency";
  const children = [
    ...(cell.cellRangeSource ? [cellRangeSourceXml(cell.cellRangeSource)] : []),
    ...decorations,
    ...(cell.detective ? [tableDetectiveXml(cell.detective)] : []),
    ...(cached.type === "string"
      ? [xmlElement("text:p", undefined, [escapeText(String(cached.value ?? ""))])]
      : []),
    ...(cell.graphics ?? []).map((graphic) => cellGraphicXml(graphic)),
  ];
  return xmlElement(
    "table:table-cell",
    {
      "table:style-name": styleOptions ? addCellStyle(styleOptions, styles) : undefined,
      "table:number-columns-spanned": span?.columnsSpanned,
      "table:number-rows-spanned": span?.rowsSpanned,
      "table:content-validation-name": contentValidationName,
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
    children,
  );
}

function cellRangeSourceXml(source: CellRangeSourceOptions): string {
  return xmlElement("table:cell-range-source", {
    "table:name": source.name,
    "table:last-column-spanned": source.lastColumnSpanned,
    "table:last-row-spanned": source.lastRowSpanned,
    "xlink:type": "simple",
    "xlink:href": source.href,
    "xlink:actuate": source.actuate,
    "table:filter-name": source.filterName,
    "table:filter-options": source.filterOptions,
    "table:refresh-delay": source.refreshDelay,
  });
}

const HIGHLIGHT_DIRECTIONS: Record<DetectiveHighlightedRangeDirection, string> = {
  fromAnotherTable: "from-another-table",
  toAnotherTable: "to-another-table",
  fromSameTable: "from-same-table",
};

const OPERATION_NAMES: Record<DetectiveOperationOptions["name"], string> = {
  traceDependents: "trace-dependents",
  removeDependents: "remove-dependents",
  tracePrecedents: "trace-precedents",
  removePrecedents: "remove-precedents",
  traceErrors: "trace-errors",
};

function tableDetectiveXml(detective: TableDetectiveOptions): string {
  return xmlElement("table:detective", undefined, [
    ...(detective.highlightedRanges ?? []).map((range) =>
      xmlElement(
        "table:highlighted-range",
        "markedInvalid" in range
          ? { "table:marked-invalid": range.markedInvalid }
          : {
              "table:cell-range-address": range.cellRangeAddress,
              "table:direction": HIGHLIGHT_DIRECTIONS[range.direction],
              "table:contains-error": range.containsError,
            },
      ),
    ),
    ...(detective.operations ?? []).map((operation) =>
      xmlElement("table:operation", {
        "table:name": OPERATION_NAMES[operation.name],
        "table:index": operation.index,
      }),
    ),
  ]);
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
  graphicStyles: Map<string, GraphicStyle> = new Map(),
): WorksheetOptions {
  const comments: CommentOptions[] = [];
  const mergeCells: MergeCellOptions[] = [];
  const shapes: WorksheetOptions["shapes"] = [];
  validateTableChildren(table);
  const columns: ColumnOptions[] = [];
  let columnNumber = 1;
  for (const column of tableColumns(table)) {
    const repeat = repeatCount(column, "table:number-columns-repeated");
    const parsed = parseColumn(column, columnNumber, dimensions);
    const remainingColumns = MAX_COLUMNS - columnNumber + 1;
    if (remainingColumns > 0) {
      const visibleRepeat = Math.min(repeat, remainingColumns);
      parsed.max = columnNumber + visibleRepeat - 1;
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
    }
    columnNumber += repeat;
  }
  const parsedRows = parseRows(
    table,
    dimensions,
    cellStyles,
    graphicStyles,
    comments,
    mergeCells,
    shapes,
  );
  const tableStyle = dimensions.get(attributeString(table, "table:style-name") ?? "");
  const sheetView =
    tableStyle?.showFormulas !== undefined || tableStyle?.showColumnHeaders !== undefined
      ? {
          ...(tableStyle?.showFormulas !== undefined
            ? { showFormulas: tableStyle.showFormulas }
            : {}),
          ...(tableStyle?.showColumnHeaders !== undefined
            ? { showRowColHeaders: tableStyle.showColumnHeaders }
            : {}),
        }
      : undefined;
  const worksheetOptions: WorksheetOptions = {
    name: attributeString(table, "table:name") ?? `Sheet${index}`,
    ...(columns.length ? { columns } : {}),
    ...(sheetView ? { sheetView } : {}),
    ...(tableStyle?.hidden ? { state: "hidden" as const } : {}),
    ...(parsedRows.defaultRowHeight === undefined
      ? {}
      : { sheetFormat: { defaultRowHeight: parsedRows.defaultRowHeight } }),
    rows: parsedRows.rows,
  };
  const selection = gnumericSelection(table);
  return {
    ...worksheetOptions,
    ...(selection ? { selection } : {}),
    ...(comments.length ? { comments } : {}),
    ...(mergeCells.length ? { mergeCells } : {}),
    ...(shapes?.length ? { shapes } : {}),
  };
}

function gnumericSelection(table: Element): SelectionOptions[] | undefined {
  const container = childNamed(table, "gnm:selections");
  if (!container) return undefined;
  const cursorColumn = attributeNumber(container, "gnm:cursor-col") ?? 0;
  const cursorRow = attributeNumber(container, "gnm:cursor-row") ?? 0;
  const ranges = childrenNamed(container, "gnm:selection").map((selection) => {
    const startColumn = (attributeNumber(selection, "gnm:start-col") ?? cursorColumn) + 1;
    const startRow = (attributeNumber(selection, "gnm:start-row") ?? cursorRow) + 1;
    const endColumn = (attributeNumber(selection, "gnm:end-col") ?? startColumn - 1) + 1;
    const endRow = (attributeNumber(selection, "gnm:end-row") ?? startRow - 1) + 1;
    return cellAddress(startColumn, startRow) === cellAddress(endColumn, endRow)
      ? cellAddress(startColumn, startRow)
      : `${cellAddress(startColumn, startRow)}:${cellAddress(endColumn, endRow)}`;
  });
  return [
    {
      activeCell: cellAddress(cursorColumn + 1, cursorRow + 1),
      ...(ranges.length ? { sqref: ranges.join(" ") } : {}),
    },
  ];
}

type ParsedSelectionRange = {
  startColumn: number;
  startRow: number;
  endColumn: number;
  endRow: number;
};

function gnumericSelectionXml(selections: readonly SelectionOptions[]): string[] {
  const selection = selections[0];
  if (!selection) return [];
  const active = parseCellReference(selection.activeCell);
  const ranges: (ParsedSelectionRange | undefined)[] = (
    selection.sqref ??
    selection.activeCell ??
    ""
  )
    .split(/\s+/u)
    .filter(Boolean)
    .map((range) => {
      const [start, end = start] = range.split(":") as [string, string?];
      const startCell = parseCellReference(start);
      const endCell = parseCellReference(end);
      if (!startCell || !endCell) return undefined;
      return {
        startColumn: startCell.column,
        startRow: startCell.row,
        endColumn: endCell.column,
        endRow: endCell.row,
      };
    });
  const firstRange = ranges[0];
  const cursor =
    active ??
    (firstRange
      ? {
          column: firstRange.startColumn,
          row: firstRange.startRow,
        }
      : undefined);
  if (!cursor) return [];
  return [
    xmlElement(
      "gnm:selections",
      {
        "gnm:cursor-col": cursor.column - 1,
        "gnm:cursor-row": cursor.row - 1,
      },
      ranges
        .filter((range): range is ParsedSelectionRange => range !== undefined)
        .map((range) =>
          xmlElement("gnm:selection", {
            "gnm:start-col": range.startColumn - 1,
            "gnm:start-row": range.startRow - 1,
            "gnm:end-col": range.endColumn - 1,
            "gnm:end-row": range.endRow - 1,
          }),
        ),
    ),
  ];
}

function parseCellReference(
  reference: string | undefined,
): { column: number; row: number } | undefined {
  const match = /^([A-Z]+)(\d+)$/u.exec(reference ?? "");
  if (!match) return undefined;
  return { column: columnNumber(match[1]!), row: Number(match[2]!) };
}

export function validateTableChildren(table: Element): void {
  for (const child of table.elements ?? []) {
    if (child.type !== "element" || !child.name) continue;
    if (
      child.name === "table:table-column" ||
      child.name === "table:table-row" ||
      child.name === "office:forms" ||
      child.name === "table:shapes" ||
      child.name === "table:named-expressions" ||
      child.name === "loext:table-protection" ||
      child.name === "table:table-source" ||
      child.name === "table:scenario" ||
      child.name === "calcext:conditional-formats" ||
      child.name === "calcext:sparkline-groups" ||
      child.name === "table:detective" ||
      child.name === "gnm:selections" ||
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
  graphicStyles: Map<string, GraphicStyle>,
  comments: CommentOptions[],
  mergeCells: MergeCellOptions[],
  shapes: WorksheetOptions["shapes"],
): { rows: RowOptions[]; defaultRowHeight?: number } {
  const rows: RowOptions[] = [];
  let defaultRowHeight: number | undefined;
  let rowNumber = 1;
  for (const row of tableRows(table)) {
    const repeat = repeatCount(row, "table:number-rows-repeated");
    const parsed = parseRow(
      row,
      rowNumber,
      dimensions,
      cellStyles,
      graphicStyles,
      comments,
      mergeCells,
      shapes,
    );
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
  graphicStyles: Map<string, GraphicStyle>,
  comments: CommentOptions[],
  mergeCells: MergeCellOptions[] = [],
  shapes: WorksheetOptions["shapes"] = [],
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
            [
              "table:style-name",
              "office:value-type",
              "calcext:value-type",
              "table:number-columns-repeated",
              "table:content-validation-name",
              "table:formula",
              "office:value",
              "office:string-value",
              "office:currency",
              "office:date-value",
              "office:boolean-value",
              "office:time-value",
            ],
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
              const updateNotification = childNamed(child, "loext:p");
              const notifyOnUpdateOfRanges = attributeString(
                updateNotification,
                "draw:notify-on-update-of-ranges",
              );
              if (updateNotification)
                assertElementNames(updateNotification, ["draw:notify-on-update-of-ranges"]);
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
                  "loext:p",
                  "xmlns:presentation",
                  "style:rel-width",
                  "style:rel-height",
                ],
                ["loext:p", "draw:image", "draw:object", "draw:object-ole", "svg:desc"],
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
                      ...(notifyOnUpdateOfRanges === undefined ? {} : { notifyOnUpdateOfRanges }),
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
                      href: isExternalOdfUrl(href) ? href : href.replace(/^\//, ""),
                      data: "",
                      ...(attributeString(child, "draw:name")
                        ? { name: attributeString(child, "draw:name") }
                        : {}),
                      ...(description ? { description } : {}),
                      ...(notifyOnUpdateOfRanges === undefined ? {} : { notifyOnUpdateOfRanges }),
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
            if (
              child.name === "draw:rect" ||
              child.name === "draw:control" ||
              child.name === "draw:ellipse" ||
              child.name === "draw:circle" ||
              child.name === "draw:line" ||
              child.name === "draw:measure" ||
              child.name === "draw:polygon" ||
              child.name === "draw:polyline" ||
              child.name === "draw:custom-shape"
            ) {
              shapes?.push(parseDrawingShape(child, rowNumber, column, graphicStyles));
            }
          }
          return parsed;
        });
      })
      .filter((cell): cell is CellOptions => cell !== undefined),
  };
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
    columnSpan - 1 > Number.MAX_SAFE_INTEGER - column ||
    rowSpan - 1 > Number.MAX_SAFE_INTEGER - row
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
      ![
        "text:p",
        "loext:p",
        "office:annotation",
        "table:cell-range-source",
        "table:detective",
        "draw:frame",
        "draw:rect",
        "draw:control",
        "draw:ellipse",
        "draw:circle",
        "draw:line",
        "draw:measure",
        "draw:polygon",
        "draw:polyline",
        "draw:custom-shape",
      ].includes(child.name ?? "")
    ) {
      throw unknownOdsElement(
        child,
        `/office:document-content/office:body/office:spreadsheet/table:table/table:table-row[${row}]/table:table-cell[${column}]`,
        child.name ?? "",
        "element has no canonical CellOptions mapping",
      );
    }
  }
  const source = childNamed(cell, "table:cell-range-source");
  if (source) result.cellRangeSource = parseCellRangeSource(source);
  const detective = childNamed(cell, "table:detective");
  if (detective) result.detective = parseTableDetective(detective);
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

function parseCellRangeSource(source: Element): CellRangeSourceOptions {
  assertElementNames(source, [
    "table:name",
    "table:last-column-spanned",
    "table:last-row-spanned",
    "xlink:type",
    "xlink:href",
    "xlink:actuate",
    "table:filter-name",
    "table:filter-options",
    "table:refresh-delay",
  ]);
  const name = attributeString(source, "table:name");
  const lastColumnSpanned = attributeNumber(source, "table:last-column-spanned");
  const lastRowSpanned = attributeNumber(source, "table:last-row-spanned");
  const href = attributeString(source, "xlink:href");
  if (
    !name ||
    !lastColumnSpanned ||
    !lastRowSpanned ||
    href === undefined ||
    attributeString(source, "xlink:type") !== "simple"
  ) {
    throw unknownOdsElement(
      source,
      "office:body/office:spreadsheet/table:table-cell",
      source.name ?? "",
      "invalid linked cell range",
    );
  }
  return {
    name,
    lastColumnSpanned,
    lastRowSpanned,
    href,
    ...(attributeString(source, "xlink:actuate") === "onRequest"
      ? { actuate: "onRequest" as const }
      : {}),
    ...(attributeString(source, "table:filter-name")
      ? { filterName: attributeString(source, "table:filter-name") }
      : {}),
    ...(attributeString(source, "table:filter-options")
      ? { filterOptions: attributeString(source, "table:filter-options") }
      : {}),
    ...(attributeString(source, "table:refresh-delay")
      ? { refreshDelay: attributeString(source, "table:refresh-delay") }
      : {}),
  };
}

const HIGHLIGHT_DIRECTION_VALUES: Record<string, DetectiveHighlightedRangeDirection> = {
  "from-another-table": "fromAnotherTable",
  "to-another-table": "toAnotherTable",
  "from-same-table": "fromSameTable",
};

const OPERATION_NAME_VALUES: Record<string, DetectiveOperationOptions["name"]> = {
  "trace-dependents": "traceDependents",
  "remove-dependents": "removeDependents",
  "trace-precedents": "tracePrecedents",
  "remove-precedents": "removePrecedents",
  "trace-errors": "traceErrors",
};

function parseTableDetective(detective: Element): TableDetectiveOptions {
  assertElementNames(detective, [], ["table:highlighted-range", "table:operation"]);
  const highlightedRanges = childrenNamed(detective, "table:highlighted-range").map(
    (range): DetectiveHighlightedRangeOptions => {
      assertElementNames(range, [
        "table:cell-range-address",
        "table:direction",
        "table:contains-error",
        "table:marked-invalid",
      ]);
      const direction = HIGHLIGHT_DIRECTION_VALUES[attributeString(range, "table:direction") ?? ""];
      const markedInvalid = attributeString(range, "table:marked-invalid");
      if (markedInvalid !== undefined) {
        if (
          direction ||
          markedInvalid === undefined ||
          !["true", "false"].includes(markedInvalid)
        ) {
          throw unknownOdsElement(
            range,
            "office:body/office:spreadsheet/table:table-cell/table:detective",
            range.name ?? "",
            "invalid highlighted range",
          );
        }
        return { markedInvalid: markedInvalid === "true" };
      }
      if (!direction) {
        throw unknownOdsElement(
          range,
          "office:body/office:spreadsheet/table:table-cell/table:detective",
          range.name ?? "",
          "highlight direction is required",
        );
      }
      const containsError = attributeString(range, "table:contains-error");
      if (containsError !== undefined && !["true", "false"].includes(containsError)) {
        throw unknownOdsElement(
          range,
          "office:body/office:spreadsheet/table:table-cell/table:detective",
          range.name ?? "",
          "invalid contains-error value",
        );
      }
      return {
        ...(attributeString(range, "table:cell-range-address")
          ? { cellRangeAddress: attributeString(range, "table:cell-range-address") }
          : {}),
        direction,
        ...(containsError !== undefined ? { containsError: containsError === "true" } : {}),
      };
    },
  );
  const operations = childrenNamed(detective, "table:operation").map(
    (operation): DetectiveOperationOptions => {
      assertElementNames(operation, ["table:name", "table:index"]);
      const name = OPERATION_NAME_VALUES[attributeString(operation, "table:name") ?? ""];
      const index = attributeNumber(operation, "table:index");
      if (!name || index === undefined || !Number.isInteger(index) || index < 0) {
        throw unknownOdsElement(
          operation,
          "office:body/office:spreadsheet/table:table-cell/table:detective",
          operation.name ?? "",
          "invalid detective operation",
        );
      }
      return { name, index };
    },
  );
  return {
    ...(highlightedRanges.length ? { highlightedRanges } : {}),
    ...(operations.length ? { operations } : {}),
  };
}

export function columnName(column: number): string {
  return spreadsheetColumnName(column);
}
