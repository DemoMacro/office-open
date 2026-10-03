import type {
  CellOptions,
  ColumnOptions,
  RowOptions,
  WorkbookOptions,
  WorksheetOptions,
} from "@office-open/xlsx";
import type { Element } from "@office-open/xml";

import { escapeText, metaXml, parseMeta } from "./meta";
import { generateOcf, readOcf, readXml, type OdfFiles } from "./package";
import {
  attributeNumber,
  attributeString,
  childNamed,
  childrenNamed,
  textOf,
  xmlElement,
} from "./xml";

const MIME = "application/vnd.oasis.opendocument.spreadsheet";
const NAMESPACES = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
].join(" ");

interface DimensionStyle {
  width?: number;
  height?: number;
  hidden?: boolean;
}

export function generateOds(options: WorkbookOptions): Uint8Array {
  const styles: string[] = [];
  const sheets = (options.worksheets ?? []).map((worksheet, index) =>
    worksheetXml(worksheet, index + 1, styles),
  );
  const files: OdfFiles = {
    "content.xml": contentXml(sheets.join(""), styles),
    "styles.xml": stylesXml(),
    "meta.xml": metaXml(options),
  };
  return generateOcf(MIME, files);
}

export function parseOds(data: Uint8Array): WorkbookOptions {
  const { files } = readOcf(data, MIME);
  const content = readXml(files, "content.xml");
  const body = childNamed(childNamed(content, "office:body"), "office:spreadsheet");
  const dimensions = parseDimensionStyles(childNamed(content, "office:automatic-styles"));
  return {
    ...parseMeta(files),
    worksheets: childrenNamed(body, "table:table").map((table, index) =>
      worksheet(table, index + 1, dimensions),
    ),
  };
}

function contentXml(sheets: string, styles: string[]): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.3"><office:automatic-styles>${styles.join(
    "",
  )}</office:automatic-styles><office:body><office:spreadsheet>${sheets}</office:spreadsheet></office:body></office:document-content>`;
}

function stylesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NAMESPACES} office:version="1.3"><office:styles/></office:document-styles>`;
}

function worksheetXml(worksheet: WorksheetOptions, index: number, styles: string[]): string {
  const columns = (worksheet.columns ?? []).flatMap((column) => {
    const count = Math.max(1, (column.max ?? column.min) - column.min + 1);
    const styleName =
      column.width === undefined && !column.hidden
        ? undefined
        : addDimensionStyle({ width: column.width, hidden: column.hidden }, styles, "column");
    return Array.from({ length: count }, () =>
      xmlElement("table:table-column", { "table:style-name": styleName }),
    );
  });
  const rows = (worksheet.rows ?? []).map((row) => {
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
      (row.cells ?? []).map((cell) => cellXml(cell)),
    );
  });
  return xmlElement("table:table", { "table:name": worksheet.name ?? `Sheet${index}` }, [
    columns.join(""),
    rows.join(""),
  ]);
}

function cellXml(cell: CellOptions): string {
  const formula = typeof cell.formula === "string" ? cell.formula : cell.formula?.formula;
  const cached = cacheAttributes(cell.value);
  return xmlElement(
    "table:table-cell",
    {
      "table:formula": formula ? `of:=${formula}` : undefined,
      "office:value-type": cached.type,
      "office:value": cached.type === "float" ? cached.value : undefined,
      "office:boolean-value": cached.type === "boolean" ? cached.value : undefined,
    },
    cached.type === "string"
      ? [xmlElement("text:p", undefined, [escapeText(String(cached.value ?? ""))])]
      : [],
  );
}

function cacheAttributes(value: CellOptions["value"]): {
  type?: string;
  value?: string | number | boolean;
} {
  if (typeof value === "number") return { type: "float", value };
  if (typeof value === "boolean") return { type: "boolean", value };
  if (typeof value === "string") return { type: "string", value };
  return {};
}

function addDimensionStyle(
  style: DimensionStyle,
  styles: string[],
  family: "column" | "row",
): string {
  const name = `${family === "column" ? "co" : "ro"}${styles.length + 1}`;
  const properties =
    family === "column"
      ? xmlElement("style:table-column-properties", {
          "style:column-width": `${style.width ?? 0}px`,
          "style:use-optimal-column-width": style.hidden ? undefined : "true",
        })
      : xmlElement("style:table-row-properties", {
          "style:row-height": style.height ? `${style.height}pt` : undefined,
        });
  styles.push(
    xmlElement("style:style", { "style:name": name, "style:family": `table-${family}` }, [
      properties,
    ]),
  );
  return name;
}

function parseDimensionStyles(container: Element | undefined): Map<string, DimensionStyle> {
  const result = new Map<string, DimensionStyle>();
  for (const style of childrenNamed(container, "style:style")) {
    const name = attributeString(style, "style:name") ?? "";
    const column = childNamed(style, "style:table-column-properties");
    const row = childNamed(style, "style:table-row-properties");
    const width = attributeString(column, "style:column-width");
    const height = attributeString(row, "style:row-height");
    result.set(name, {
      width: width?.endsWith("px") ? Number(width.slice(0, -2)) : undefined,
      height: height?.endsWith("pt") ? Number(height.slice(0, -2)) : undefined,
      hidden:
        attributeString(column, "style:column-hidden") === "true" ||
        attributeString(row, "style:row-hidden") === "true" ||
        undefined,
    });
  }
  return result;
}

function worksheet(
  table: Element,
  index: number,
  dimensions: Map<string, DimensionStyle>,
): WorksheetOptions {
  return {
    name: attributeString(table, "table:name") ?? `Sheet${index}`,
    columns: parseColumns(table, dimensions),
    rows: childrenNamed(table, "table:table-row").map((row, rowIndex) =>
      parseRow(row, rowIndex + 1, dimensions),
    ),
  };
}

function parseColumns(table: Element, dimensions: Map<string, DimensionStyle>): ColumnOptions[] {
  const columns: ColumnOptions[] = [];
  let column = 1;
  for (const element of childrenNamed(table, "table:table-column")) {
    const style = dimensions.get(attributeString(element, "table:style-name") ?? "");
    columns.push({ min: column, max: column, width: style?.width, hidden: style?.hidden });
    column += 1;
  }
  return columns;
}

function parseRow(
  row: Element,
  rowNumber: number,
  dimensions: Map<string, DimensionStyle>,
): RowOptions {
  const style = dimensions.get(attributeString(row, "table:style-name") ?? "");
  return {
    rowNumber,
    height: style?.height,
    hidden: style?.hidden,
    cells:
      row.elements
        ?.filter((cell) => cell.name === "table:table-cell")
        .map((cell, cellIndex) => parseCell(cell, rowNumber, cellIndex + 1)) ?? [],
  };
}

function parseCell(cell: Element, row: number, column: number): CellOptions {
  const result: CellOptions = { reference: `${columnName(column)}${row}` };
  const formula = attributeString(cell, "table:formula");
  if (formula) result.formula = formula.replace(/^of:=/, "");
  const valueType = attributeString(cell, "office:value-type");
  if (valueType === "float") result.value = attributeNumber(cell, "office:value");
  else if (valueType === "boolean")
    result.value = attributeString(cell, "office:boolean-value") === "true";
  else if (valueType === "string") result.value = textOf(childNamed(cell, "text:p"));
  return result;
}

function columnName(column: number): string {
  let result = "";
  let value = column;
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}
