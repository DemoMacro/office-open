import type {
  CellOptions,
  ColumnOptions,
  RowOptions,
  WorkbookOptions,
  WorksheetOptions,
} from "@office-open/xlsx";
import type { Element } from "@office-open/xml";

import { escapeText, metaXml, parseMeta } from "./meta";
import { parseOdfNodes, serializeOdfNodes, type OdfXmlNode } from "./odf-node";
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
  'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
].join(" ");

interface DimensionStyle {
  width?: number;
  height?: number;
  hidden?: boolean;
}

export type OdsOptions = WorkbookOptions & { odfExtensions?: OdfXmlNode[] };

export function generateOds(options: OdsOptions): Uint8Array {
  const styles: string[] = [];
  const sheets = (options.worksheets ?? []).map((worksheet, index) =>
    worksheetXml(worksheet, index + 1, styles),
  );
  const files: OdfFiles = {
    "content.xml": contentXml(
      [...sheets, ...serializeOdfNodes(options.odfExtensions)].join(""),
      styles,
    ),
    "styles.xml": stylesXml(),
    "meta.xml": metaXml(options),
  };
  return generateOcf(MIME, files);
}

export function parseOds(data: Uint8Array): OdsOptions {
  const { files } = readOcf(data, MIME);
  const content = readXml(files, "content.xml");
  const body = childNamed(childNamed(content, "office:body"), "office:spreadsheet");
  const automaticStyles = childNamed(content, "office:automatic-styles");
  const dimensions = parseDimensionStyles(automaticStyles);
  const cellStyles = parseNumberStyles(automaticStyles);
  const rawNodes = parseOdfNodes(body);
  return {
    ...parseMeta(files),
    worksheets: childrenNamed(body, "table:table").map((table, index) =>
      worksheet(table, index + 1, dimensions, cellStyles),
    ),
    odfExtensions: rawNodes.filter((node) => node.name !== "table:table"),
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
      (row.cells ?? []).map((cell) => cellXml(cell, styles)),
    );
  });
  const rows = bodyRows.length
    ? bodyRows
    : [xmlElement("table:table-row", undefined, [xmlElement("table:table-cell")])];
  return xmlElement("table:table", { "table:name": worksheet.name ?? `Sheet${index}` }, [
    columns.join(""),
    rows.join(""),
  ]);
}

function cellXml(cell: CellOptions, styles: string[]): string {
  const formula = typeof cell.formula === "string" ? cell.formula : cell.formula?.formula;
  const cached = cacheAttributes(cell.value);
  const numFmt = typeof cell.style === "object" ? cell.style.numFmt : undefined;
  return xmlElement(
    "table:table-cell",
    {
      "table:style-name": numFmt ? addNumberStyle(numFmt, styles) : undefined,
      "table:formula": formula ? `of:=${formula}` : undefined,
      "office:value-type": cached.type,
      "office:value": cached.type === "float" ? cached.value : undefined,
      "office:boolean-value": cached.type === "boolean" ? cached.value : undefined,
      "office:date-value": cached.type === "date" ? cached.value : undefined,
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
  if (value instanceof Date) return { type: "date", value: odfDateValue(value) };
  if (typeof value === "string") return { type: "string", value };
  return {};
}

/** Registers a numFmt as a data style + table-cell style; returns the cell style name. */
function addNumberStyle(numFmt: string, styles: string[]): string {
  const data = numFmtDataStyle(numFmt);
  const dataName = `N${styles.length + 1}`;
  styles.push(xmlElement(data.name, { "style:name": dataName }, data.children));
  const styleName = `ce${styles.length + 1}`;
  styles.push(
    xmlElement("style:style", {
      "style:name": styleName,
      "style:family": "table-cell",
      "style:data-style-name": dataName,
    }),
  );
  return styleName;
}

/** Translates an Excel numFmt code into the matching ODF data style element. */
function numFmtDataStyle(numFmt: string): { name: string; children: string[] } {
  const percent = numFmt.endsWith("%");
  const core = percent ? numFmt.slice(0, -1) : numFmt;
  const tokens = core.match(/yyyy|yy|mm|dd|hh|ss|:/g);
  if (tokens && tokens.join("") === core) {
    const isTime = tokens.includes("hh") || tokens.includes("ss");
    if (isTime) {
      return {
        name: "number:time-style",
        children: tokens.map((token) => {
          const long = { hh: "number:hours", mm: "number:minutes", ss: "number:seconds" }[token];
          return long
            ? xmlElement(long, { "number:style": "long" })
            : xmlElement("number:text", undefined, [token]);
        }),
      };
    }
    return {
      name: "number:date-style",
      children: tokens.map((token) => {
        const element = {
          yyyy: "number:year",
          yy: "number:year",
          mm: "number:month",
          dd: "number:day",
        }[token];
        return element
          ? xmlElement(element, {
              "number:style": token.length === 2 && token !== "yy" ? "long" : undefined,
            })
          : xmlElement("number:text", undefined, [token]);
      }),
    };
  }
  const grouping = core.includes(",");
  const decimals = core.split(".")[1];
  if (!/^[#0,]*(\.[0]+)?$/.test(core)) {
    return {
      name: "number:number-style",
      children: [xmlElement("number:text", undefined, [numFmt])],
    };
  }
  const digits = Math.max(1, (core.split(".")[0] ?? "").replace(/[#,]/g, "").length);
  const children = [
    xmlElement("number:number", {
      "number:decimal-places": decimals?.length,
      "number:min-integer-digits": digits,
      "number:grouping": grouping || undefined,
    }),
  ];
  if (percent) children.push(xmlElement("number:text", undefined, ["%"]));
  return {
    name: percent ? "number:percentage-style" : "number:number-style",
    children,
  };
}

/** ODF office:date-value is UTC ISO 8601 without timezone suffix. */
function odfDateValue(date: Date): string {
  return date.toISOString().slice(0, 19);
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

/** Maps table-cell style names back to their numFmt codes via data styles. */
function parseNumberStyles(container: Element | undefined): Map<string, string> {
  const dataStyles = new Map<string, string>();
  const cellStyles = new Map<string, string>();
  for (const element of container?.elements ?? []) {
    if (
      element.name === "number:number-style" ||
      element.name === "number:percentage-style" ||
      element.name === "number:date-style" ||
      element.name === "number:time-style"
    ) {
      const format = dataStyleNumFmt(element);
      if (format) dataStyles.set(attributeString(element, "style:name") ?? "", format);
    }
  }
  for (const style of childrenNamed(container, "style:style")) {
    if (attributeString(style, "style:family") !== "table-cell") continue;
    const dataName = attributeString(style, "style:data-style-name");
    const format = dataName ? dataStyles.get(dataName) : undefined;
    if (format) cellStyles.set(attributeString(style, "style:name") ?? "", format);
  }
  return cellStyles;
}

/** Reconstructs an Excel numFmt code from ODF number:* style children. */
function dataStyleNumFmt(element: Element): string | undefined {
  const parts: string[] = [];
  for (const child of element.elements ?? []) {
    if (child.type === "text") {
      parts.push(String(child.text ?? ""));
      continue;
    }
    const long = attributeString(child, "number:style") === "long";
    switch (child.name) {
      case "number:number": {
        const decimals = attributeNumber(child, "number:decimal-places") ?? 0;
        const digits = attributeNumber(child, "number:min-integer-digits") ?? 1;
        parts.push(
          (attributeString(child, "number:grouping") === "true" ? "#,##0" : "0".repeat(digits)) +
            (decimals > 0 ? `.${"0".repeat(decimals)}` : ""),
        );
        break;
      }
      case "number:year":
        parts.push(long ? "yyyy" : "yy");
        break;
      case "number:month":
        parts.push(long ? "mm" : "m");
        break;
      case "number:day":
        parts.push(long ? "dd" : "d");
        break;
      case "number:hours":
        parts.push(long ? "hh" : "h");
        break;
      case "number:minutes":
        parts.push(long ? "mm" : "m");
        break;
      case "number:seconds":
        parts.push(long ? "ss" : "s");
        break;
      case "number:text":
        parts.push(textOf(child));
        break;
      default:
        return undefined;
    }
  }
  let format = parts.join("");
  if (element.name === "number:percentage-style" && !format.endsWith("%")) format += "%";
  return format;
}

function worksheet(
  table: Element,
  index: number,
  dimensions: Map<string, DimensionStyle>,
  cellStyles: Map<string, string>,
): WorksheetOptions {
  return {
    name: attributeString(table, "table:name") ?? `Sheet${index}`,
    columns: parseColumns(table, dimensions),
    rows: childrenNamed(table, "table:table-row").map((row, rowIndex) =>
      parseRow(row, rowIndex + 1, dimensions, cellStyles),
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
  cellStyles: Map<string, string>,
): RowOptions {
  const style = dimensions.get(attributeString(row, "table:style-name") ?? "");
  return {
    rowNumber,
    height: style?.height,
    hidden: style?.hidden,
    cells:
      row.elements
        ?.filter((cell) => cell.name === "table:table-cell")
        .map((cell, cellIndex) => parseCell(cell, rowNumber, cellIndex + 1, cellStyles)) ?? [],
  };
}

function parseCell(
  cell: Element,
  row: number,
  column: number,
  cellStyles: Map<string, string>,
): CellOptions {
  const result: CellOptions = { reference: `${columnName(column)}${row}` };
  const numFmt = cellStyles.get(attributeString(cell, "table:style-name") ?? "");
  if (numFmt) result.style = { numFmt };
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
