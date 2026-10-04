import type {
  ChartSpaceOptions,
  FormContainerOptions,
  FormControlOptions,
} from "@office-open/core";
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
  type OdfPackageFiles,
  type XmlAttributes,
} from "@office-open/ocf";
import {
  CHART_MIME,
  chartBodyXml,
  OdfSchemaError,
  parseEmbeddedCharts,
} from "@office-open/odf-schema";
import type {
  AlignmentOptions,
  BorderOptions,
  BorderSideOptions,
  CellFillOptions,
  CellOptions,
  ColumnOptions,
  CommentOptions,
  FontOptions,
  RowOptions,
  StyleOptions,
  WorkbookOptions,
  WorksheetChartOptions,
  WorksheetOptions,
} from "@office-open/xlsx";
import type { Element } from "@office-open/xml";

import { OdsParseError } from "./error";

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

type OdsParsedChartOptions = WorksheetChartOptions;

export function generateOds(options: WorkbookOptions): Uint8Array {
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
  const files: OdfPackageFiles = {
    "content.xml": contentXml(
      sheets.join(""),
      styles,
      options.definedNames,
      options.calculation,
      options.forms,
    ),
    "styles.xml": stylesXml(),
    "meta.xml": metaXml(options),
  };
  for (const entry of chartFrames) files[`${entry.name}/content.xml`] = chartBodyXml(entry.chart);
  for (const worksheet of options.worksheets ?? []) {
    for (const row of worksheet.rows ?? []) {
      for (const cell of row.cells ?? []) {
        for (const graphic of cell.graphics ?? []) {
          if (graphic.type === "image") files[graphic.href] = base64ToBytes(graphic.data);
          else {
            if (graphic.chart) files[`${graphic.href}/content.xml`] = chartBodyXml(graphic.chart);
          }
        }
      }
    }
  }
  return generateOcf(
    MIME,
    files,
    Object.fromEntries([
      ...chartFrames.map((entry) => [`${entry.name}/`, CHART_MIME]),
      ...(options.worksheets ?? []).flatMap((worksheet) =>
        (worksheet.rows ?? []).flatMap((row) =>
          (row.cells ?? []).flatMap((cell) =>
            (cell.graphics ?? []).map((graphic) => [
              graphic.type === "image" ? graphic.href : `${graphic.href}/`,
              graphic.type === "image" ? imageMediaType(graphic.href) : CHART_MIME,
            ]),
          ),
        ),
      ),
    ]),
  );
}

function base64ToBytes(value: string): Uint8Array {
  return Uint8Array.from(atob(value), (character) => character.charCodeAt(0));
}

function bytesToBase64(value: Uint8Array): string {
  let binary = "";
  for (const byte of value) binary += String.fromCharCode(byte);
  return btoa(binary);
}

function imageMediaType(href: string): string {
  const types: Record<string, string> = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    bmp: "image/bmp",
    svg: "image/svg+xml",
  };
  return types[href.split(".").pop() ?? ""] ?? "application/octet-stream";
}

function unsupportedOdsValue(name: string, reason: string): OdsParseError {
  return new OdsParseError(
    `content.xml: ${name} ${reason}`,
    "content.xml",
    "/office:document-content/office:body/office:spreadsheet",
    name,
    "no canonical ODS mapping",
  );
}

export function parseOds(data: Uint8Array): WorkbookOptions {
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

function parseOdsWorkbook(data: Uint8Array): WorkbookOptions {
  const { files, binaries, manifest } = readOcf(data, MIME);
  const content = readXml(files, "content.xml");
  const body = childNamed(childNamed(content, "office:body"), "office:spreadsheet");
  const automaticStyles = childNamed(content, "office:automatic-styles");
  const dimensions = parseDimensionStyles(automaticStyles);
  const cellStyles = parseNumberStyles(automaticStyles);
  const chartPool = parseEmbeddedCharts(manifest, files);
  const calcProperties = parseOdsSemantics(body);
  const forms = parseForms(body);
  const embeddedCharts = [...parseWorksheetCharts(body, chartPool)];
  const definedNames = parseDefinedNames(body) ?? [];
  const worksheets = childrenNamed(body, "table:table").map((table, index) => {
    const parsed = worksheet(table, index + 1, dimensions, cellStyles);
    const name = attributeString(table, "table:name");
    const charts = embeddedCharts
      .filter((entry) => entry.worksheet === (name ?? `Sheet${index + 1}`))
      .map((entry) => entry.chart);
    return charts.length > 0
      ? { ...parsed, charts: [...(parsed.charts ?? []), ...charts] }
      : parsed;
  });
  for (const worksheet of worksheets) {
    for (const row of worksheet.rows ?? []) {
      for (const cell of row.cells ?? []) {
        for (const graphic of cell.graphics ?? []) {
          if (graphic.type === "image") {
            const binary = binaries[graphic.href];
            graphic.data = binary ? bytesToBase64(binary) : "";
            if (!graphic.data)
              throw unsupportedOdsValue(graphic.href, "referenced image is missing");
          } else {
            const chart = chartPool.get(graphic.href);
            if (chart) graphic.chart = chart;
          }
        }
      }
    }
  }
  rejectUnknownSpreadsheetChildren(body);
  const result = {
    ...parseMeta(files),
    ...(definedNames.length > 0 ? { definedNames } : {}),
    ...(calcProperties ? { calculation: calcProperties } : {}),
    ...(forms.length > 0 ? { forms } : {}),
    worksheets,
  };
  return result;
}

function contentXml(
  sheets: string,
  styles: string[],
  definedNames: WorkbookOptions["definedNames"],
  calculation?: WorkbookOptions["calculation"],
  forms?: WorkbookOptions["forms"],
): string {
  const expressions = `${formsXml(forms)}${semanticsXml(calculation)}${definedNamesXml(
    definedNames,
  )}`;
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.3"><office:automatic-styles>${styles.join(
    "",
  )}</office:automatic-styles><office:body><office:spreadsheet>${sheets}${expressions}</office:spreadsheet></office:body></office:document-content>`;
}

function parseOdsSemantics(body: Element | undefined): WorkbookOptions["calculation"] {
  const calculation = childNamed(body, "table:calculation-settings");
  if (!calculation) return undefined;
  const iteration = childNamed(calculation, "table:iteration");
  const result: NonNullable<WorkbookOptions["calculation"]> = {
    ...(iteration
      ? {
          iterate: attributeString(iteration, "table:status") === "enable",
          ...(attributeNumber(iteration, "table:steps") !== undefined
            ? { iterateCount: attributeNumber(iteration, "table:steps") }
            : {}),
          ...(attributeNumber(iteration, "table:maximum-difference") !== undefined
            ? { iterateDelta: attributeNumber(iteration, "table:maximum-difference") }
            : {}),
        }
      : {}),
    ...(attributeString(calculation, "table:case-sensitive") === "false"
      ? { caseSensitive: false }
      : {}),
    ...(attributeString(calculation, "table:precision-as-shown") === "true"
      ? { precisionAsShown: true }
      : {}),
    ...(attributeString(calculation, "table:automatic-find-labels") === "true"
      ? { automaticLabelSearch: true }
      : {}),
    ...(attributeString(calculation, "table:use-regular-expressions") === "true"
      ? { regularExpressionSearch: true }
      : {}),
    ...(attributeString(calculation, "table:use-wildcards") === "true"
      ? { wildcardSearch: true }
      : {}),
  };
  const nullDate = childNamed(calculation, "table:null-date");
  if (nullDate) {
    result.nullDate = {
      year: attributeNumber(nullDate, "table:date-value-year") ?? 1899,
      month: attributeNumber(nullDate, "table:date-value-month") ?? 12,
      day: attributeNumber(nullDate, "table:date-value-day") ?? 30,
    };
  }
  assertCalculationChildren(calculation, result);
  return result;
}

function assertCalculationChildren(
  calculation: Element,
  result: NonNullable<WorkbookOptions["calculation"]>,
): void {
  for (const child of calculation.elements ?? []) {
    if (child.type !== "element") continue;
    if (child.name === "table:iteration") {
      assertElementNames(child, ["table:status", "table:steps", "table:maximum-difference"]);
      continue;
    }
    if (child.name === "table:null-date") {
      assertElementNames(child, [
        "table:date-value-year",
        "table:date-value-month",
        "table:date-value-day",
      ]);
      if (!result.nullDate) result.nullDate = { year: 1899, month: 12, day: 30 };
      continue;
    }
    throw unknownOdsElement(
      child,
      "/office:document-content/office:body/office:spreadsheet/table:calculation-settings",
      child.name ?? "",
      "element has no canonical calculation mapping",
    );
  }
}

function formsXml(forms: WorkbookOptions["forms"]): string {
  if (!forms?.length) return "";
  return xmlElement(
    "office:forms",
    {
      "form:automatic-focus": forms[0]?.automaticFocus,
      "form:apply-design-mode": forms[0]?.designMode,
    },
    forms.map((form) =>
      xmlElement(
        "form:form",
        { "form:name": form.name },
        form.controls.map((control) => formControlXml(control)),
      ),
    ),
  );
}

function formControlXml(control: FormControlOptions): string {
  const common = {
    "form:id": control.id,
    "form:name": control.name,
    "form:disabled": control.disabled,
    "form:tab-index": control.tabIndex,
    "form:automatic-focus": control.automaticFocus,
  };
  if (control.control === "checkBox")
    return xmlElement("form:checkbox", {
      ...common,
      "form:current-state": control.checked ? "checked" : "unchecked",
    });
  if (control.control === "dropDownList")
    return xmlElement(
      "form:listbox",
      common,
      control.entries.map((label, optionIndex) =>
        xmlElement(
          "form:option",
          { "form:selected": control.selectedIndex === optionIndex || undefined },
          [escapeText(label)],
        ),
      ),
    );
  return xmlElement("form:text", {
    ...common,
    "office:value-type": control.valueType,
    "form:current-value": control.value,
    "form:max-length": control.maxLength,
  });
}

function parseForms(body: Element | undefined): FormContainerOptions[] {
  return descendantElements(body)
    .filter((element) => element.name === "office:forms")
    .flatMap((element) => {
      assertElementNames(
        element,
        ["form:automatic-focus", "form:apply-design-mode"],
        ["form:form"],
      );
      return childrenNamed(element, "form:form").map((form) => {
        assertElementNames(form, ["form:name"], ["form:text", "form:checkbox", "form:listbox"]);
        return {
          ...(attributeString(form, "form:name")
            ? { name: attributeString(form, "form:name") }
            : {}),
          ...(attributeString(element, "form:automatic-focus") === "true"
            ? { automaticFocus: true }
            : {}),
          ...(attributeString(element, "form:apply-design-mode") === "true"
            ? { designMode: true }
            : {}),
          controls:
            form.elements
              ?.filter((child) => child.type === "element")
              .map((control) => parseFormControl(control)) ?? [],
        };
      });
    });
}

function descendantElements(element: Element | undefined): Element[] {
  return (element?.elements ?? []).flatMap((child) =>
    child.type === "element" ? [child, ...descendantElements(child)] : [],
  );
}

function parseFormControl(control: Element): FormControlOptions {
  assertElementNames(
    control,
    [
      "form:id",
      "form:name",
      "form:disabled",
      "form:tab-index",
      "form:automatic-focus",
      "office:value-type",
      "form:current-value",
      "form:max-length",
      "form:current-state",
      "form:selected",
    ],
    ["form:option", "form:item"],
  );
  const common = {
    ...(attributeString(control, "form:id") ? { id: attributeString(control, "form:id") } : {}),
    ...(attributeString(control, "form:name")
      ? { name: attributeString(control, "form:name") }
      : {}),
    ...(attributeString(control, "form:disabled") === "true" ? { disabled: true } : {}),
    ...(attributeNumber(control, "form:tab-index") !== undefined
      ? { tabIndex: attributeNumber(control, "form:tab-index") }
      : {}),
    ...(attributeString(control, "form:automatic-focus") === "true"
      ? { automaticFocus: true }
      : {}),
  };
  if (control.name === "form:text")
    return {
      control: "text",
      ...common,
      ...(attributeString(control, "form:current-value")
        ? { value: attributeString(control, "form:current-value") }
        : {}),
      ...(attributeString(control, "office:value-type")
        ? { valueType: attributeString(control, "office:value-type") as "float" | "string" }
        : {}),
      ...(attributeNumber(control, "form:max-length") !== undefined
        ? { maxLength: attributeNumber(control, "form:max-length") }
        : {}),
    };
  if (control.name === "form:checkbox") {
    const state = attributeString(control, "form:current-state");
    if (state !== "checked" && state !== "unchecked")
      throw unknownOdsElement(control, "", control.name ?? "", "invalid checkbox state");
    return { control: "checkBox", ...common, checked: state === "checked" };
  }
  if (control.name === "form:listbox") {
    const entries = childrenNamed(control, "form:option").map((option) => textOf(option));
    const selected = childrenNamed(control, "form:option").findIndex(
      (option) => attributeString(option, "form:selected") === "true",
    );
    return {
      control: "dropDownList",
      ...common,
      entries,
      ...(selected >= 0 ? { selectedIndex: selected } : {}),
    };
  }
  throw unknownOdsElement(
    control,
    "/office:spreadsheet/office:forms",
    control.name ?? "",
    "element has no canonical control mapping",
  );
}

function semanticsXml(calculation: WorkbookOptions["calculation"]): string {
  if (!calculation) return "";
  const hasSettings =
    calculation.iterate !== undefined ||
    calculation.caseSensitive !== undefined ||
    calculation.precisionAsShown !== undefined ||
    calculation.automaticLabelSearch !== undefined ||
    calculation.regularExpressionSearch !== undefined ||
    calculation.wildcardSearch !== undefined ||
    calculation.nullDate !== undefined;
  if (!hasSettings) return "";
  const iteration = {
    "table:status":
      calculation.iterate === undefined ? undefined : calculation.iterate ? "enable" : "disable",
    "table:steps": calculation.iterateCount,
    "table:maximum-difference": calculation.iterateDelta,
  };
  return xmlElement(
    "table:calculation-settings",
    {
      "table:case-sensitive": calculation.caseSensitive,
      "table:precision-as-shown": calculation.precisionAsShown,
      "table:automatic-find-labels": calculation.automaticLabelSearch,
      "table:use-regular-expressions": calculation.regularExpressionSearch,
      "table:use-wildcards": calculation.wildcardSearch,
    },
    [
      ...(calculation.iterate !== undefined ? [xmlElement("table:iteration", iteration)] : []),
      ...(calculation.nullDate
        ? [
            xmlElement("table:null-date", {
              "table:date-value-year": calculation.nullDate.year,
              "table:date-value-month": calculation.nullDate.month,
              "table:date-value-day": calculation.nullDate.day,
            }),
          ]
        : []),
    ],
  );
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
          col: 1,
          row: 1,
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

function assertElementNames(
  element: Element,
  allowedAttributes: string[],
  allowedChildren: string[] = [],
): void {
  const unknown = Object.keys(element.attributes ?? {}).filter(
    (name) => !allowedAttributes.includes(name),
  );
  if (unknown.length) {
    throw unknownOdsElement(
      element,
      `/office:document-content/office:body${element.name === "form:form" ? "/office:spreadsheet/office:forms" : ""}`,
      unknown[0] ?? "",
      "attribute has no canonical mapping",
    );
  }
  for (const child of element.elements ?? []) {
    if (child.type === "element" && !allowedChildren.includes(child.name ?? "")) {
      throw unknownOdsElement(
        element,
        `/office:document-content/office:body${element.name === "form:form" ? "/office:spreadsheet/office:forms" : ""}`,
        child.name ?? "",
        "child has no canonical mapping",
      );
    }
  }
}

function cellXml(cell: CellOptions, styles: string[], decorations: string[] = []): string {
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

function cellGraphicXml(graphic: NonNullable<CellOptions["graphics"]>[number]): string {
  const common = {
    "draw:name": graphic.name,
    "svg:x": emuToLength(graphic.x),
    "svg:y": emuToLength(graphic.y),
    "svg:width": emuToLength(graphic.width),
    "svg:height": emuToLength(graphic.height),
    "draw:z-index": graphic.zIndex,
    "table:end-cell-address": graphic.endCellAddress,
    ...(graphic.endX !== undefined ? { "table:end-x": emuToLength(graphic.endX) } : {}),
    ...(graphic.endY !== undefined ? { "table:end-y": emuToLength(graphic.endY) } : {}),
    "draw:id": graphic.id,
    "draw:style-name": graphic.styleName,
    "draw:text-style-name": graphic.textStyleName,
    "style:rel-width": graphic.relativeWidth,
    "style:rel-height": graphic.relativeHeight,
  };
  const link = {
    "xlink:href": graphic.href,
    "xlink:type": "simple",
    "xlink:show": "embed",
    "xlink:actuate": "onLoad",
  };
  return xmlElement("draw:frame", common, [
    xmlElement(graphic.type === "image" ? "draw:image" : "draw:object", link),
    ...(graphic.description
      ? [xmlElement("svg:desc", undefined, [escapeText(graphic.description)])]
      : []),
  ]);
}

function commentPlainText(text: CommentOptions["text"]): string {
  if (typeof text === "string") return text;
  return text.runs?.map((run) => run.text).join("") ?? text.text ?? "";
}

function cellDecorationXml(worksheet: WorksheetOptions, reference: string): string[] {
  const comment = worksheet.comments?.find((entry) => entry.cell === reference);
  return comment
    ? [
        xmlElement(
          "office:annotation",
          {
            "dc:creator": comment.author,
          },
          [xmlElement("text:p", undefined, [escapeText(commentPlainText(comment.text))])],
        ),
      ]
    : [];
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

function validateTableChildren(table: Element): void {
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
