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
  RowOptions,
  StyleOptions,
  WorksheetOptions,
} from "@office-open/xlsx";
import type { Element } from "@office-open/xml";

import { cellDecorationXml } from "./comments";
import { cellGraphicXml, type OdsChartFrame } from "./drawing-chart";
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

export function worksheetXml(
  worksheet: WorksheetOptions,
  index: number,
  styles: string[],
  embeddedCharts: OdsChartFrame[],
): string {
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
  const maxRowCells = Math.max(0, ...(worksheet.rows ?? []).map((row) => row.cells?.length ?? 0));
  const columnCount = Math.max(explicitColumns.length, maxRowCells, 1);
  const columns = Array.from(
    { length: columnCount },
    (_, column) => explicitColumns[column] ?? xmlElement("table:table-column"),
  );
  const bodyRows = (worksheet.rows ?? []).map((row) => {
    const styleName =
      row.height === undefined && !row.hidden
        ? undefined
        : addDimensionStyle(
            { height: typeof row.height === "number" ? row.height : undefined, hidden: row.hidden },
            styles,
            "row",
          );
    return xmlElement(
      "table:table-row",
      { "table:style-name": styleName },
      (row.cells ?? []).map((cell) =>
        cellXml(cell, styles, cellDecorationXml(worksheet, cell.reference ?? "")),
      ),
    );
  });
  const rows = bodyRows.length
    ? bodyRows
    : [xmlElement("table:table-row", undefined, [xmlElement("table:table-cell")])];
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
    columns.join(""),
    rows.join(""),
  ]);
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

export function cellXml(cell: CellOptions, styles: string[], decorations: string[] = []): string {
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
  return xmlElement(
    "table:table-cell",
    {
      "table:style-name": styleOptions ? addCellStyle(styleOptions, styles) : undefined,
      "table:formula": formula ? `of:=${formula}` : undefined,
      "office:value-type": cached.type,
      "office:value": cached.type === "float" ? cached.value : undefined,
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

export function worksheet(
  table: Element,
  index: number,
  dimensions: Map<string, DimensionStyle>,
  cellStyles: Map<string, StyleOptions>,
): WorksheetOptions {
  const comments: CommentOptions[] = [];
  validateTableChildren(table);
  const worksheetOptions = {
    name: attributeString(table, "table:name") ?? `Sheet${index}`,
    columns: tableColumns(table).map((column, columnIndex) =>
      parseColumn(column, columnIndex + 1, dimensions),
    ),
    rows: tableRows(table).map((row, rowIndex) =>
      parseRow(row, rowIndex + 1, dimensions, cellStyles, comments),
    ),
  };
  return {
    ...worksheetOptions,
    ...(comments.length ? { comments } : {}),
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

export function parseRow(
  row: Element,
  rowNumber: number,
  dimensions: Map<string, DimensionStyle>,
  cellStyles: Map<string, StyleOptions>,
  comments: CommentOptions[],
): RowOptions {
  const style = dimensions.get(attributeString(row, "table:style-name") ?? "");
  return {
    rowNumber,
    height: style?.height,
    hidden: style?.hidden,
    cells: (row.elements ?? [])
      .flatMap((cell, cellIndex) => {
        const column = cellIndex + 1;
        if (cell.type !== "element" || !cell.name) return undefined;
        if (cell.name === "table:covered-table-cell") {
          assertElementNames(
            cell,
            ["table:style-name", "office:value-type", "calcext:value-type"],
            ["text:p"],
          );
          return [
            {
              reference: `${columnName(column)}${rowNumber}`,
              covered: {
                reference: `${columnName(column)}${rowNumber}`,
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
            },
          ];
        }
        if (cell.name !== "table:table-cell") {
          throw unknownOdsElement(
            cell,
            `/office:document-content/office:body/office:spreadsheet/table:table/table:table-row[${rowNumber}]`,
            cell.name,
            "element has no canonical CellOptions mapping",
          );
        }
        const parsed = parseCell(cell, rowNumber, column, cellStyles);
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
            const object = childNamed(child, "draw:object") ?? childNamed(child, "draw:object-ole");
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
                    href: href.replace(/^\//, ""),
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
        return [parsed];
      })
      .filter((cell): cell is CellOptions => cell !== undefined),
  };
}

export function parseCell(
  cell: Element,
  row: number,
  column: number,
  cellStyles: Map<string, StyleOptions>,
): CellOptions {
  const result: CellOptions = { reference: `${columnName(column)}${row}` };
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
  else if (valueType === "boolean")
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
