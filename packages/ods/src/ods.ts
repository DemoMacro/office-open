import type { ChartSpaceOptions } from "@office-open/core";
import {
  attributeNumber,
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  escapeText,
  generateOcf,
  lengthToEmu,
  metaXml,
  parseMeta,
  readOcf,
  readXml,
  textOf,
  xmlElement,
  type OdfFiles,
  type XmlAttributes,
} from "@office-open/ocf";
import {
  CHART_MIME,
  chartBodyXml,
  officeFormsXml,
  OdfSchemaError,
  parseEmbeddedCharts,
  parseOfficeForms,
} from "@office-open/odf-schema";
import type {
  AlignmentOptions,
  BorderOptions,
  BorderSideOptions,
  CellFillOptions,
  CellOptions,
  ColumnOptions,
  FontOptions,
  RowOptions,
  StyleOptions,
  WorkbookOptions,
  WorksheetChartOptions,
} from "@office-open/xlsx";
import type { Element } from "@office-open/xml";

import { OdsParseError } from "./error";
import type {
  OdsAnnotation,
  OdsCellGraphic,
  OdsCellOptions,
  OdsObjectGraphic,
  OdsSemanticsOptions,
  OdsWorkbookOptions,
  OdsWorksheetOptions,
} from "./semantics";

const MIME = "application/vnd.oasis.opendocument.spreadsheet";
const NAMESPACES = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
  'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"',
  'xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"',
  'xmlns:xforms="http://www.w3.org/2002/xforms"',
  'xmlns:dc="http://purl.org/dc/elements/1.1/"',
].join(" ");

const COLUMN_CONTAINERS = new Set([
  "table:table-columns",
  "table:table-header-columns",
  "table:table-column-group",
]);
const ROW_CONTAINERS = new Set([
  "table:table-rows",
  "table:table-header-rows",
  "table:table-row-group",
]);

interface DimensionStyle {
  width?: number;
  height?: number;
  hidden?: boolean;
}

/** A worksheet chart frame collected while serializing an ODS body. */
interface OdsChartFrame {
  name: string;
  worksheet: string;
  chart: OdsParsedChartOptions;
}

type OdsParsedChartOptions = Omit<WorksheetChartOptions, "col" | "row"> &
  Partial<Pick<WorksheetChartOptions, "col" | "row">>;

export function generateOds(options: OdsWorkbookOptions): Uint8Array {
  const styles: string[] = [];
  const chartFrames = (options.worksheets ?? []).flatMap((worksheet, worksheetIndex) =>
    (worksheet.charts ?? []).map((chart, chartIndex) => ({
      name: chart.name ?? `Object ${worksheetIndex + chartIndex + 1}`,
      worksheet: worksheet.name ?? `Sheet${worksheetIndex + 1}`,
      chart,
    })),
  );
  const sheets = (options.worksheets ?? []).map((worksheet, index) =>
    worksheetXml(worksheet, index + 1, styles, chartFrames),
  );
  const files: OdfFiles = {
    "content.xml": contentXml(sheets.join(""), styles, options.definedNames, options.odfSemantics),
    "styles.xml": stylesXml(),
    "meta.xml": metaXml(options),
  };
  for (const entry of chartFrames) files[`${entry.name}/content.xml`] = chartBodyXml(entry.chart);
  return generateOcf(
    MIME,
    files,
    Object.fromEntries(chartFrames.map((entry) => [`${entry.name}/`, CHART_MIME])),
  );
}

export function parseOds(data: Uint8Array): OdsWorkbookOptions {
  try {
    return parseOdsWorkbook(data);
  } catch (cause) {
    if (cause instanceof OdsParseError) throw cause;
    if (cause instanceof OdfSchemaError) {
      throw new OdsParseError(cause.message, cause.part, cause.path, cause.name, cause.reason, {
        cause,
      });
    }
    throw new OdsParseError(
      cause instanceof Error ? cause.message : "Unable to parse ODS package",
      "mimetype",
      "/mimetype",
      "mimetype",
      "invalid ODS package",
      { cause },
    );
  }
}

function parseOdsWorkbook(data: Uint8Array): OdsWorkbookOptions {
  const { files, manifest } = readOcf(data, MIME);
  const content = readXml(files, "content.xml");
  const body = childNamed(childNamed(content, "office:body"), "office:spreadsheet");
  const automaticStyles = childNamed(content, "office:automatic-styles");
  const dimensions = parseDimensionStyles(automaticStyles);
  const cellStyles = parseNumberStyles(automaticStyles);
  const chartPool = parseEmbeddedCharts(manifest, files);
  const semantics = parseOdsSemantics(body);
  const embeddedCharts = [...parseWorksheetCharts(body, chartPool)];
  const definedNames = parseDefinedNames(body) ?? [];
  const worksheets = childrenNamed(body, "table:table").map((table, index) => {
    const parsed = worksheet(table, index + 1, dimensions, cellStyles) as OdsWorksheetOptions;
    const formsElement = childNamed(table, "office:forms");
    if (formsElement) {
      parsed.forms = parseOfficeForms(
        formsElement,
        "content.xml",
        `/office:document-content/office:body/office:spreadsheet/table:table/office:forms`,
      );
    }
    const objectGraphics = childrenNamed(childNamed(table, "table:shapes"), "draw:frame")
      .filter((frame) => {
        const object = childNamed(frame, "draw:object") ?? childNamed(frame, "draw:object-ole");
        return object !== undefined && !attributeString(object, "xlink:href");
      })
      .map((frame) => ({
        reference: "",
        href: attributeString(childNamed(frame, "draw:object"), "xlink:href"),
        name: attributeString(frame, "draw:name"),
        x: lengthToEmu(attributeString(frame, "svg:x")),
        y: lengthToEmu(attributeString(frame, "svg:y")),
        width: lengthToEmu(attributeString(frame, "svg:width")),
        height: lengthToEmu(attributeString(frame, "svg:height")),
      }));
    if (objectGraphics.length) parsed.objectGraphics = objectGraphics;
    const name = attributeString(table, "table:name");
    const charts = embeddedCharts
      .filter((entry) => entry.worksheet === (name ?? `Sheet${index + 1}`))
      .map((entry) => entry.chart);
    return charts.length > 0
      ? { ...parsed, charts: [...(parsed.charts ?? []), ...(charts as WorksheetChartOptions[])] }
      : parsed;
  });
  rejectUnknownSpreadsheetChildren(body);
  const result = {
    ...parseMeta(files),
    ...(definedNames.length > 0 ? { definedNames } : {}),
    ...(semantics ? { odfSemantics: semantics } : {}),
    worksheets,
  };
  return result;
}

function contentXml(
  sheets: string,
  styles: string[],
  definedNames: WorkbookOptions["definedNames"],
  semantics?: OdsSemanticsOptions,
): string {
  const expressions = `${semanticsXml(semantics)}${definedNamesXml(definedNames)}`;
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.3"><office:automatic-styles>${styles.join(
    "",
  )}</office:automatic-styles><office:body><office:spreadsheet>${sheets}${expressions}</office:spreadsheet></office:body></office:document-content>`;
}

function parseOdsSemantics(body: Element | undefined): OdsSemanticsOptions | undefined {
  const result: OdsSemanticsOptions = {};
  const forms = childNamed(body, "office:forms");
  if (forms)
    result.forms = parseOfficeForms(
      forms,
      "content.xml",
      "/office:document-content/office:body/office:spreadsheet/office:forms",
    );
  const calculation = childNamed(body, "table:calculation-settings");
  if (calculation) {
    const nullDate = childNamed(calculation, "table:null-date");
    const iteration = childNamed(calculation, "table:iteration");
    result.calculationSettings = {
      caseSensitive: attributeString(calculation, "table:case-sensitive") !== "false",
      automaticFindLabels: attributeString(calculation, "table:automatic-find-labels") === "true",
      regularExpressions: attributeString(calculation, "table:use-regular-expressions") === "true",
      wildcards: attributeString(calculation, "table:use-wildcards") === "true",
      nullDate: nullDate ? attributeString(nullDate, "table:date-value") : undefined,
      iteration: iteration
        ? {
            enabled: attributeString(iteration, "table:status") === "enable",
            steps: attributeNumber(iteration, "table:steps"),
            maximumDifference: attributeNumber(iteration, "table:maximum-difference"),
          }
        : undefined,
    };
  }
  return Object.keys(result).length ? result : undefined;
}

function semanticsXml(semantics: OdsSemanticsOptions | undefined): string {
  if (!semantics) return "";
  const calculation = semantics.calculationSettings;
  return [
    officeFormsXml(semantics.forms),
    calculation
      ? xmlElement(
          "table:calculation-settings",
          {
            "table:case-sensitive": calculation.caseSensitive === false ? false : undefined,
            "table:automatic-find-labels": calculation.automaticFindLabels,
            "table:use-regular-expressions": calculation.regularExpressions,
            "table:use-wildcards": calculation.wildcards,
          },
          [
            ...(calculation.nullDate
              ? [xmlElement("table:null-date", { "table:date-value": calculation.nullDate })]
              : []),
            ...(calculation.iteration
              ? [
                  xmlElement("table:iteration", {
                    "table:status": calculation.iteration.enabled ? "enable" : "disable",
                    "table:steps": calculation.iteration.steps,
                    "table:maximum-difference": calculation.iteration.maximumDifference,
                  }),
                ]
              : []),
          ],
        )
      : "",
  ].join("");
}

function definedNamesXml(definedNames: WorkbookOptions["definedNames"]): string {
  if (!definedNames?.length) return "";
  return xmlElement(
    "table:named-expressions",
    undefined,
    definedNames.map((definedName) =>
      /^[A-Za-z0-9_]+![A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?$/.test(definedName.value)
        ? xmlElement("table:named-range", {
            "table:name": definedName.name,
            "table:cell-range-address": definedName.value,
          })
        : xmlElement("table:named-expression", {
            "table:name": definedName.name,
            "table:expression": definedName.value,
          }),
    ),
  );
}

function stylesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NAMESPACES} office:version="1.3"><office:styles/></office:document-styles>`;
}

function worksheetXml(
  worksheet: OdsWorksheetOptions,
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
        (cell as OdsCellOptions).covered
          ? xmlElement("table:covered-table-cell")
          : cellXml(cell, styles, cellDecorationXml(worksheet, cell.reference ?? "")),
      ),
    );
  });
  const rows = bodyRows.length
    ? bodyRows
    : [xmlElement("table:table-row", undefined, [xmlElement("table:table-cell")])];
  const sheetName = worksheet.name ?? `Sheet${index}`;
  const objectFrames = (worksheet.objectGraphics ?? []).map((object) =>
    xmlElement(
      "draw:frame",
      {
        "draw:name": object.name,
        "svg:x": emuToLength(object.x ?? 0),
        "svg:y": emuToLength(object.y ?? 0),
        "svg:width": emuToLength(object.width ?? 0),
        "svg:height": emuToLength(object.height ?? 0),
      },
      [
        object.href
          ? xmlElement("draw:object", {
              "xlink:href": object.href,
              "xlink:type": "simple",
              "xlink:show": "embed",
              "xlink:actuate": "onLoad",
            })
          : xmlElement("draw:object-ole", undefined, [xmlElement("office:binary-data")]),
      ],
    ),
  );
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
    officeFormsXml(worksheet.forms),
    ...(objectFrames.length || frames.length
      ? [xmlElement("table:shapes", undefined, [...objectFrames, ...frames])]
      : []),
    columns.join(""),
    rows.join(""),
  ]);
}

/** Worksheet-anchored draw:frame chart objects resolved from the pool. */
function parseWorksheetCharts(
  body: Element | undefined,
  pool: Map<string, ChartSpaceOptions>,
): OdsChartFrame[] {
  const charts: OdsChartFrame[] = [];
  for (const table of childrenNamed(body, "table:table")) {
    const worksheet = attributeString(table, "table:name");
    const frames: Element[] = [];
    for (const child of table.elements ?? []) {
      if (child.name === "draw:frame") frames.push(child);
      if (child.name === "table:shapes") {
        for (const nested of child.elements ?? [])
          if (nested.name === "draw:frame") frames.push(nested);
      }
    }
    for (const frame of frames) {
      const href = attributeString(childNamed(frame, "draw:object"), "xlink:href")
        ?.replace(/^\.\//, "")
        .replace(/\/$/, "");
      const chart = href ? pool.get(href) : undefined;
      if (!href) continue;
      if (!chart) {
        throw unknownOdsElement(
          frame,
          `/office:spreadsheet/table:table[@table:name="${worksheet ?? ""}"]`,
          href,
          "referenced chart subdocument is missing",
        );
      }
      charts.push({
        name: href,
        worksheet: worksheet ?? "",
        chart: {
          ...chart,
          name: href,
          anchorType: "absolute",
          ...endAnchor(frame),
          absoluteX: lengthToEmu(attributeString(frame, "svg:x")),
          absoluteY: lengthToEmu(attributeString(frame, "svg:y")),
          extentCx: lengthToEmu(attributeString(frame, "svg:width")) ?? 400000,
          extentCy: lengthToEmu(attributeString(frame, "svg:height")) ?? 300000,
        },
      });
    }
  }
  return charts;
}

function endAnchor(frame: Element): Pick<WorksheetChartOptions, "toCol" | "toRow"> {
  const address = attributeString(frame, "table:end-cell-address");
  const match = /^([A-Z]+)([0-9]+)$/.exec(address ?? "");
  if (!match) return {};
  return {
    toCol: columnNumber(match[1]!),
    toRow: Number(match[2]),
  };
}

function cellAddress(column: number, row: number): string {
  let value = column;
  let letters = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    value = Math.floor((value - 1) / 26);
  }
  return `${letters}${row}`;
}

function columnNumber(column: string): number {
  let total = 0;
  for (let index = 0; index < column.length; index += 1) {
    total = total * 26 + column.charCodeAt(index) - 64;
  }
  return total;
}

function rejectUnknownSpreadsheetChildren(body: Element | undefined): void {
  const allowed = new Set([
    "table:table",
    "table:calculation-settings",
    "office:forms",
    "table:named-expressions",
    "draw:frame",
  ]);
  for (const child of body?.elements ?? []) {
    if (child.name && !allowed.has(child.name)) {
      const name = child.name;
      throw new OdsParseError(
        `content.xml: /office:document-content/office:body/office:spreadsheet/${name}: no canonical WorkbookOptions mapping`,
        "content.xml",
        `/office:document-content/office:body/office:spreadsheet/${name}`,
        name,
        "no canonical WorkbookOptions mapping",
      );
    }
  }
}

function parseDefinedNames(body: Element | undefined): WorkbookOptions["definedNames"] {
  const expressions = childNamed(body, "table:named-expressions");
  return [
    ...childrenNamed(expressions, "table:named-range").map((range) => ({
      name: attributeString(range, "table:name") ?? "",
      value: attributeString(range, "table:cell-range-address") ?? "",
    })),
    ...childrenNamed(expressions, "table:named-expression").map((expression) => ({
      name: attributeString(expression, "table:name") ?? "",
      value: attributeString(expression, "table:expression") ?? "",
    })),
  ];
}

function unknownOdsElement(
  element: Element,
  parent: string,
  name: string,
  reason: string,
): OdsParseError {
  const path = `${parent}/${element.name ?? name}`;
  return new OdsParseError(
    `content.xml: ${path}: ${name}: ${reason}`,
    "content.xml",
    path,
    name,
    reason,
  );
}

function cellXml(cell: CellOptions, styles: string[], decorations: string[] = []): string {
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
    ],
  );
}

function cellDecorationXml(worksheet: OdsWorksheetOptions, reference: string): string[] {
  const annotation = worksheet.annotations?.find((entry) => entry.reference === reference);
  const graphic = worksheet.cellGraphics?.find((entry) => entry.reference === reference);
  return [
    ...(annotation
      ? [
          xmlElement(
            "office:annotation",
            {
              "dc:creator": annotation.author,
              "dc:date": annotation.date,
            },
            annotation.paragraphs.map((paragraph) =>
              xmlElement("text:p", undefined, [escapeText(paragraph)]),
            ),
          ),
        ]
      : []),
    ...(graphic
      ? [
          xmlElement(
            "draw:frame",
            {
              "draw:name": graphic.name,
              "svg:x": graphic.x !== undefined ? emuToLength(graphic.x) : undefined,
              "svg:y": graphic.y !== undefined ? emuToLength(graphic.y) : undefined,
              "svg:width": graphic.width !== undefined ? emuToLength(graphic.width) : undefined,
              "svg:height": graphic.height !== undefined ? emuToLength(graphic.height) : undefined,
            },
            [xmlElement("draw:image", { "xlink:href": graphic.href, "xlink:type": "simple" })],
          ),
        ]
      : []),
  ];
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

/** Registers a table-cell style (numFmt data style + font/fill); returns its name. */
function addCellStyle(style: StyleOptions, styles: string[]): string {
  const data = style.numFmt ? numFmtDataStyle(style.numFmt) : undefined;
  const dataName = `N${styles.length + 1}`;
  if (data) styles.push(xmlElement(data.name, { "style:name": dataName }, data.children));
  const styleName = `ce${styles.length + 1}`;
  styles.push(
    xmlElement(
      "style:style",
      {
        "style:name": styleName,
        "style:family": "table-cell",
        "style:data-style-name": data ? dataName : undefined,
      },
      cellStyleChildren(style),
    ),
  );
  return styleName;
}

/** Font/fill properties as ODF style children; empty array when styleless. */
function cellStyleChildren(style: StyleOptions): string[] {
  const font = style.font;
  const fill = style.fill;
  const alignment = style.alignment;
  const textAttributes = {
    "fo:color": odfHex(font?.color),
    "fo:font-size": font?.size !== undefined ? `${font.size}pt` : undefined,
    "fo:font-weight": font?.bold ? "bold" : undefined,
    "fo:font-style": font?.italic ? "italic" : undefined,
    "fo:underline-style": font?.underline ? "solid" : undefined,
    "fo:text-line-through-style": font?.strike ? "solid" : undefined,
  };
  const cellAttributes = {
    "fo:background-color": odfHex(solidFillColor(fill)),
    "style:vertical-align": odfVertical(alignment?.vertical),
    "fo:wrap-option": alignment?.wrapText ? "wrap" : undefined,
    ...borderAttributes(style.border),
  };
  const paragraphAttributes = { "fo:text-align": odfHorizontal(alignment?.horizontal) };
  const children: string[] = [];
  if (Object.values(cellAttributes).some((value) => value !== undefined))
    children.push(xmlElement("style:table-cell-properties", cellAttributes));
  if (Object.values(paragraphAttributes).some((value) => value !== undefined))
    children.push(xmlElement("style:paragraph-properties", paragraphAttributes));
  if (Object.values(textAttributes).some((value) => value !== undefined))
    children.push(xmlElement("style:text-properties", textAttributes));
  return children;
}

/** XLSX cell borders → ODF border shorthand with explicit RGB colors. */
function borderAttributes(border: BorderSideOptions | undefined): XmlAttributes {
  const attributes = {
    "fo:border-top": odfBorder(border?.top),
    "fo:border-bottom": odfBorder(border?.bottom),
    "fo:border-left": odfBorder(border?.left),
    "fo:border-right": odfBorder(border?.right),
    // xlsx diagonal flags map onto ODF's two named diagonal directions.
    "style:diagonal-tl-br": border?.diagonalDown ? odfBorder(border.diagonal) : undefined,
    "style:diagonal-bl-tr": border?.diagonalUp ? odfBorder(border.diagonal) : undefined,
  };
  return Object.values(attributes).some((value) => value !== undefined) ? attributes : {};
}

/** XLSX border side → CSS-style ODF shorthand (`width style color`). */
function odfBorder(side: BorderOptions | undefined): string | undefined {
  if (!side || side.style === undefined) return undefined;
  if (side.style === "none") return "none";
  const width = side.style === "thick" ? "2.5pt" : mediumBorder(side.style) ? "1pt" : "0.5pt";
  const style = odfLineStyle(side.style);
  return [width, style, odfHex(side.color)].filter(Boolean).join(" ");
}

function mediumBorder(style: BorderOptions["style"]): boolean {
  return (
    style === "medium" ||
    style === "mediumDashed" ||
    style === "mediumDashDot" ||
    style === "mediumDashDotDot"
  );
}

/** Closest ODF/CSS line style; compound XLSX dashes become dashed. */
function odfLineStyle(style: NonNullable<BorderOptions["style"]>): string {
  if (
    style === "dashed" ||
    style === "mediumDashed" ||
    style === "dashDot" ||
    style === "mediumDashDot"
  )
    return "dashed";
  if (style === "dotted" || style === "dashDotDot" || style === "mediumDashDotDot") return "dotted";
  if (style === "double") return "double";
  return "solid";
}

/** ODF border shorthand → the closest typed XLSX border side. */
function parseBorder(value: string | undefined): BorderOptions | undefined {
  if (!value || value === "none") return value === "none" ? { style: "none" } : undefined;
  const match = /^(-?\d+(?:\.\d+)?(?:cm|mm|pt|pc|in|px))\s+(\S+)(?:\s+#([0-9a-fA-F]{6}))?$/.exec(
    value,
  );
  if (!match) return undefined;
  const [, width, lineStyle, color] = match;
  const points = lengthToEmu(width)! / 12700;
  const style = ((): BorderOptions["style"] => {
    if (lineStyle === "dashed") return points >= 0.75 ? "mediumDashed" : "dashed";
    if (lineStyle === "dotted") return "dotted";
    if (lineStyle === "double") return "double";
    if (points >= 1.75) return "thick";
    if (points >= 0.75) return "medium";
    if (points < 0.25) return "hair";
    return "thin";
  })();
  return { style, color: color?.toUpperCase() };
}

/** xlsx horizontal → ODF fo:text-align. */
function odfHorizontal(value: AlignmentOptions["horizontal"]): string | undefined {
  if (value === "left") return "start";
  if (value === "right") return "end";
  if (value === "center" || value === "justify") return value;
  return undefined;
}

/** xlsx vertical → ODF style:vertical-align. */
function odfVertical(value: AlignmentOptions["vertical"]): string | undefined {
  if (value === "center") return "middle";
  if (value === "top" || value === "bottom") return value;
  return undefined;
}

/** Solid-fill foreground color; pattern/gradient fills keep their typed shape. */
function solidFillColor(fill: CellFillOptions | undefined): string | undefined {
  return fill?.type === undefined || fill.type === "solid" ? fill?.color : undefined;
}

/** xlsx hex (RRGGBB or AARRGGBB) → ODF #RRGGBB. */
function odfHex(hex: string | undefined): string | undefined {
  return hex ? `#${hex.slice(-6)}` : undefined;
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

/** Maps table-cell style names back to typed StyleOptions via data styles. */
function parseNumberStyles(container: Element | undefined): Map<string, StyleOptions> {
  const dataStyles = new Map<string, string>();
  const cellStyles = new Map<string, StyleOptions>();
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
    const options: StyleOptions = {};
    const dataName = attributeString(style, "style:data-style-name");
    const numFmt = dataName ? dataStyles.get(dataName) : undefined;
    if (numFmt) options.numFmt = numFmt;
    const text = childNamed(style, "style:text-properties");
    const size = attributeString(text, "fo:font-size");
    const font: FontOptions = {
      color: odfColor(attributeString(text, "fo:color")),
      size: size?.endsWith("pt") ? Number(size.slice(0, -2)) : undefined,
      bold: attributeString(text, "fo:font-weight") === "bold" || undefined,
      italic: attributeString(text, "fo:font-style") === "italic" || undefined,
      underline: attributeString(text, "fo:underline-style") === "solid" || undefined,
      strike: attributeString(text, "fo:text-line-through-style") === "solid" || undefined,
    };
    if (Object.values(font).some((value) => value !== undefined)) options.font = font;
    const background = odfColor(
      attributeString(childNamed(style, "style:table-cell-properties"), "fo:background-color"),
    );
    if (background) options.fill = { type: "solid", color: background };
    const cellProps = childNamed(style, "style:table-cell-properties");
    const border: BorderSideOptions = {
      top: parseBorder(attributeString(cellProps, "fo:border-top")),
      bottom: parseBorder(attributeString(cellProps, "fo:border-bottom")),
      left: parseBorder(attributeString(cellProps, "fo:border-left")),
      right: parseBorder(attributeString(cellProps, "fo:border-right")),
    };
    const diagonalDown = parseBorder(attributeString(cellProps, "style:diagonal-tl-br"));
    const diagonalUp = parseBorder(attributeString(cellProps, "style:diagonal-bl-tr"));
    if (diagonalDown !== undefined || diagonalUp !== undefined) {
      border.diagonal = diagonalDown ?? diagonalUp;
      border.diagonalDown = diagonalDown !== undefined || undefined;
      border.diagonalUp = diagonalUp !== undefined || undefined;
    }
    if (Object.values(border).some((side) => side !== undefined)) options.border = border;
    const vertical = attributeString(cellProps, "style:vertical-align");
    const wrapText = attributeString(cellProps, "fo:wrap-option") === "wrap";
    const horizontal = attributeString(
      childNamed(style, "style:paragraph-properties"),
      "fo:text-align",
    );
    const alignment: AlignmentOptions = {
      horizontal:
        horizontal === "start"
          ? "left"
          : horizontal === "end"
            ? "right"
            : (horizontal as AlignmentOptions["horizontal"]),
      vertical: (vertical === "middle" ? "center" : vertical) as AlignmentOptions["vertical"],
      wrapText: wrapText || undefined,
    };
    if (Object.values(alignment).some((value) => value !== undefined))
      options.alignment = alignment;
    if (Object.keys(options).length > 0) {
      cellStyles.set(attributeString(style, "style:name") ?? "", options);
    }
  }
  return cellStyles;
}

/** ODF #RRGGBB → xlsx RRGGBB hex. */
function odfColor(value: string | undefined): string | undefined {
  return value?.startsWith("#") ? value.slice(1) : undefined;
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
  cellStyles: Map<string, StyleOptions>,
): OdsWorksheetOptions {
  const annotations: OdsAnnotation[] = [];
  const cellGraphics: OdsCellGraphic[] = [];
  const objectGraphics: OdsObjectGraphic[] = [];
  validateTableChildren(table);
  const worksheetOptions = {
    name: attributeString(table, "table:name") ?? `Sheet${index}`,
    columns: tableColumns(table).map((column, columnIndex) =>
      parseColumn(column, columnIndex + 1, dimensions),
    ),
    rows: tableRows(table).map((row, rowIndex) =>
      parseRow(
        row,
        rowIndex + 1,
        dimensions,
        cellStyles,
        annotations,
        cellGraphics,
        objectGraphics,
      ),
    ),
  };
  return {
    ...worksheetOptions,
    ...(annotations.length ? { annotations } : {}),
    ...(cellGraphics.length ? { cellGraphics } : {}),
    ...(objectGraphics.length ? { objectGraphics } : {}),
  };
}

function validateTableChildren(table: Element): void {
  for (const child of table.elements ?? []) {
    if (child.type !== "element" || !child.name) continue;
    if (
      child.name === "table:table-column" ||
      child.name === "table:table-row" ||
      child.name === "table:shapes" ||
      child.name === "office:forms" ||
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

function tableColumns(table: Element): Element[] {
  return (table.elements ?? []).flatMap((child) => {
    if (child.name === "table:table-column") return [child];
    if (COLUMN_CONTAINERS.has(child.name ?? "")) return tableColumns(child);
    return [];
  });
}

function tableRows(table: Element): Element[] {
  return (table.elements ?? []).flatMap((child) => {
    if (child.name === "table:table-row") return [child];
    if (ROW_CONTAINERS.has(child.name ?? "")) return tableRows(child);
    return [];
  });
}

function parseColumn(
  element: Element,
  column: number,
  dimensions: Map<string, DimensionStyle>,
): ColumnOptions {
  const style = dimensions.get(attributeString(element, "table:style-name") ?? "");
  return { min: column, max: column, width: style?.width, hidden: style?.hidden };
}

function parseRow(
  row: Element,
  rowNumber: number,
  dimensions: Map<string, DimensionStyle>,
  cellStyles: Map<string, StyleOptions>,
  annotations: OdsAnnotation[],
  cellGraphics: OdsCellGraphic[],
  objectGraphics: OdsObjectGraphic[],
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
          return {
            reference: `${columnName(column)}${rowNumber}`,
            covered: true,
          } as OdsCellOptions;
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
            annotations.push({
              reference: parsed.reference ?? "",
              paragraphs: childrenNamed(child, "text:p").map((paragraph) => textOf(paragraph)),
              author: attributeString(child, "dc:creator"),
              date: attributeString(child, "dc:date"),
            });
          }
          if (child.name === "draw:frame") {
            const object = childNamed(child, "draw:object");
            const image = childNamed(child, "draw:image");
            const href = attributeString(object ?? image, "xlink:href");
            if (!href) {
              objectGraphics.push({
                reference: parsed.reference ?? "",
                name: attributeString(child, "draw:name"),
                x: lengthToEmu(attributeString(child, "svg:x")),
                y: lengthToEmu(attributeString(child, "svg:y")),
                width: lengthToEmu(attributeString(child, "svg:width")),
                height: lengthToEmu(attributeString(child, "svg:height")),
              });
              continue;
            }
            cellGraphics.push({
              reference: parsed.reference ?? "",
              href,
              name: attributeString(child, "draw:name"),
              x: lengthToEmu(attributeString(child, "svg:x")),
              y: lengthToEmu(attributeString(child, "svg:y")),
              width: lengthToEmu(attributeString(child, "svg:width")),
              height: lengthToEmu(attributeString(child, "svg:height")),
            });
          }
        }
        return [parsed];
      })
      .filter((cell): cell is CellOptions => cell !== undefined),
  };
}

function parseCell(
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
