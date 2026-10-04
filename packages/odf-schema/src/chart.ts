import type {
  AxisOptions,
  ChartSeriesData,
  ChartSpaceOptions,
  ChartTitleOptions,
} from "@office-open/core";
import {
  attributeString,
  childNamed,
  childrenNamed,
  escapeText,
  generateOcf,
  metaXml,
  ODF_NAMESPACES,
  parseMeta,
  parseOdfNodes,
  readOcf,
  readXml,
  serializeOdfNodes,
  textOf,
  xmlElement,
  type OdfFiles,
  type OdfXmlNode,
} from "@office-open/ocf";
import { parse, type Element } from "@office-open/xml";

import { OdfSchemaError } from "./error";

const MIME = "application/vnd.oasis.opendocument.chart";
export const CHART_MIME = MIME;
const NAMESPACES = `${ODF_NAMESPACES} xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"`;

/** Standalone ODF chart subdocument using the shared chart model. */
export interface ChartDocumentOptions {
  title?: string;
  chart: ChartSpaceOptions;
}

const CHART_CLASSES: Record<ChartSpaceOptions["type"], string> = {
  column: "chart:bar",
  bar: "chart:bar",
  line: "chart:line",
  pie: "chart:circle",
  area: "chart:area",
  scatter: "chart:scatter",
  bubble: "chart:scatter",
  doughnut: "chart:ring",
  radar: "chart:radar",
  stock: "chart:stock",
  surface: "chart:fill",
  ofPie: "chart:circle",
};

const CLASS_CHARTS: Record<string, ChartSpaceOptions["type"]> = {
  "chart:bar": "column",
  "chart:line": "line",
  "chart:circle": "pie",
  "chart:ring": "doughnut",
  "chart:area": "area",
  "chart:scatter": "scatter",
  "chart:radar": "radar",
  "chart:filled-radar": "radar",
  "chart:stock": "stock",
  "chart:fill": "surface",
};

export function generateChartDocument(options: ChartDocumentOptions): Uint8Array {
  return generateOcf(MIME, {
    "content.xml": chartContentXml(chartXml(options.chart)),
    "meta.xml": metaXml({ title: options.title }),
  });
}

/** Serializes an embedded chart subdocument body (its content.xml). */
export function chartBodyXml(chart: ChartSpaceOptions): string {
  return chartContentXml(chartXml(chart));
}

export function parseChartDocument(data: Uint8Array): ChartDocumentOptions {
  const { files } = readOcf(data, MIME);
  const content = readXml(files, "content.xml");
  const body = childNamed(childNamed(content, "office:body"), "office:chart");
  const chart = parseChart(childNamed(body, "chart:chart"));
  if (!chart) throw new OdfSchemaError("Chart document is missing chart:chart");
  return { ...parseMeta(files), chart };
}

/** Parses an embedded chart body; undefined when the XML is not a chart document. */
export function parseChartBody(xml: string): ChartSpaceOptions | undefined {
  const document = parse(xml, { ignoreDeclaration: true, ignoreDoctype: true });
  const root = document.elements?.[0] ?? document;
  const body = childNamed(childNamed(root, "office:body"), "office:chart");
  return parseChart(childNamed(body, "chart:chart"));
}

/** Embedded chart subdocuments from the manifest, keyed by their object name. */
export function parseEmbeddedCharts(
  manifest: Element,
  files: OdfFiles,
): Map<string, ChartSpaceOptions> {
  const result = new Map<string, ChartSpaceOptions>();
  for (const entry of childrenNamed(manifest, "manifest:file-entry")) {
    const fullPath = attributeString(entry, "manifest:full-path");
    if (!fullPath || fullPath === "/" || !fullPath.endsWith("/")) continue;
    const content = files[`${fullPath}content.xml`];
    const chart = content ? parseChartBody(content) : undefined;
    if (chart) result.set(fullPath.replace(/\/$/, ""), chart);
  }
  return result;
}

function chartContentXml(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.3"><office:body><office:chart>${body}</office:chart></office:body></office:document-content>`;
}

function chartXml(chart: ChartSpaceOptions): string {
  const local = localTable(chart);
  return xmlElement("chart:chart", { "chart:class": CHART_CLASSES[chart.type] }, [
    titleXml("chart:title", chart.title),
    chart.showLegend === false ? "" : "<chart:legend/>",
    plotAreaXml(chart, local),
    ...(local ? serializeOdfNodes([local]) : []),
  ]);
}

function titleText(title: string | ChartTitleOptions | undefined): string | undefined {
  if (typeof title === "string") return title;
  return typeof title?.text === "string" ? title.text : undefined;
}

function titleXml(name: string, title: string | ChartTitleOptions | undefined): string {
  const text = titleText(title);
  if (text === undefined) return "";
  return xmlElement(name, undefined, [xmlElement("text:p", undefined, [escapeText(text)])]);
}

function plotAreaXml(chart: ChartSpaceOptions, local: OdfXmlNode | undefined): string {
  const series = chart.series.map((item, index) => seriesXml(item, index, local, chart.type));
  return xmlElement("chart:plot-area", undefined, [
    ...(chart.axes ?? defaultAxes()).map((axis) => axisXml(axis, chart.categoryFormula, local)),
    ...series,
  ]);
}

function axisXml(
  axis: AxisOptions,
  categoryFormula: string | undefined,
  local: OdfXmlNode | undefined,
): string {
  const categories =
    axis.kind === "category"
      ? xmlElement("chart:categories", {
          "table:cell-range-address":
            categoryFormula ?? (local ? `local-table.A2:A${chartRowCount(local)}` : undefined),
        })
      : "";
  return xmlElement("chart:axis", { "chart:dimension": axis.kind === "value" ? "y" : "x" }, [
    titleXml("chart:title", axis.title),
    categories,
    axis.majorGridlines ? '<chart:grid chart:class="major"/>' : "",
    axis.minorGridlines ? '<chart:grid chart:class="minor"/>' : "",
  ]);
}

function seriesXml(
  series: ChartSpaceOptions["series"][number],
  index: number,
  local: OdfXmlNode | undefined,
  type: ChartSpaceOptions["type"],
): string {
  const column = String.fromCharCode("B".charCodeAt(0) + index);
  const values =
    series.valueFormula ??
    (local ? `local-table.${column}2:${column}${seriesValues(series).length + 1}` : undefined);
  const label = series.nameFormula ?? (local ? `local-table.${column}1` : undefined);
  return xmlElement("chart:series", {
    "chart:values-cell-range-address": values,
    "chart:label-cell-address": label,
    "chart:class": CHART_CLASSES[type],
  });
}

function defaultAxes(): AxisOptions[] {
  return [{ kind: "category" }, { kind: "value", majorGridlines: true }];
}

function chartRowCount(local: OdfXmlNode): number {
  return (local.children ?? []).filter((child) => typeof child !== "string").length;
}

function localTable(chart: ChartSpaceOptions): OdfXmlNode | undefined {
  if (chart.series.some((series) => series.valueFormula)) return undefined;
  const categories = [...(chart.categories ?? [])];
  const rows = chart.series.reduce(
    (maximum, series) => Math.max(maximum, seriesValues(series).length),
    0,
  );
  if (rows === 0) return undefined;
  const header = ["", ...chart.series.map((series, index) => series.name ?? `Series ${index + 1}`)];
  return {
    name: "table:table",
    attributes: { "table:name": "local-table" },
    children: [
      tableRow(header),
      ...Array.from({ length: rows }, (_, row) =>
        tableRow([
          categories[row] ?? "",
          ...chart.series.map((series) => String(seriesValues(series)[row] ?? "")),
        ]),
      ),
    ],
  };
}

function tableRow(values: string[]): OdfXmlNode {
  return {
    name: "table:table-row",
    children: values.map((value) => ({
      name: "table:table-cell",
      attributes: { "office:value-type": "float", "office:value": value },
      children: [value],
    })),
  };
}

function seriesValues(series: ChartSpaceOptions["series"][number]): readonly number[] {
  return "values" in series ? series.values : series.yValues;
}

function parseChart(element: Element | undefined): ChartSpaceOptions | undefined {
  if (!element) return undefined;
  const classToken = attributeString(element, "chart:class");
  const type = classToken ? CLASS_CHARTS[classToken] : undefined;
  if (!type) throw new OdfSchemaError(`Unsupported ODF chart class: ${classToken ?? "missing"}`);
  const plotArea = childNamed(element, "chart:plot-area");
  const series = childrenNamed(plotArea, "chart:series").map((child) =>
    parseSeries(child, element),
  );
  return {
    type,
    title: titleText(parseTitle(childNamed(element, "chart:title"))),
    categories: parseCategories(plotArea, series, element),
    series,
    showLegend: childNamed(element, "chart:legend") !== undefined,
    axes: childrenNamed(plotArea, "chart:axis").map(parseAxis),
  };
}

function parseTitle(element: Element | undefined): ChartTitleOptions | undefined {
  const text = textOf(childNamed(element, "text:p"));
  return text ? { text } : undefined;
}

function parseAxis(element: Element): AxisOptions {
  const dimension = attributeString(element, "chart:dimension");
  return {
    kind: dimension === "y" ? "value" : "category",
    title: titleText(parseTitle(childNamed(element, "chart:title"))),
    majorGridlines: childrenNamed(element, "chart:grid").some(
      (grid) => attributeString(grid, "chart:class") !== "minor",
    ),
    minorGridlines: childrenNamed(element, "chart:grid").some(
      (grid) => attributeString(grid, "chart:class") === "minor",
    ),
  };
}

function parseSeries(element: Element, chart: Element): ChartSeriesData {
  const valuesRange = attributeString(element, "chart:values-cell-range-address");
  const labelRange = attributeString(element, "chart:label-cell-address");
  return {
    name: labelRange ? localValue(chart, labelRange) : undefined,
    nameFormula: labelRange?.startsWith("local-table.") ? undefined : labelRange,
    valueFormula: valuesRange?.startsWith("local-table.") ? undefined : valuesRange,
    values: parseLocalTable(chart, valuesRange),
  };
}

function parseCategories(
  plotArea: Element | undefined,
  series: ChartSeriesData[],
  chart: Element,
): string[] | undefined {
  const axis = childrenNamed(plotArea, "chart:axis").find(
    (item) => attributeString(item, "chart:dimension") !== "y",
  );
  const range = attributeString(childNamed(axis, "chart:categories"), "table:cell-range-address");
  if (!range) return undefined;
  const categories = localCells(chart, range);
  return categories.length > 0 || series.length > 0 ? categories : undefined;
}

function parseLocalTable(chart: Element, range: string | undefined): number[] {
  if (!range?.startsWith("local-table.")) return [];
  return localCells(chart, range).map(Number).filter(Number.isFinite);
}

function localValue(chart: Element, range: string): string | undefined {
  return localCells(chart, range)[0];
}

function localCells(chart: Element, range: string | undefined): string[] {
  if (!range) return [];
  const match = /^local-table\.([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(range);
  if (!match) return [];
  const table = parseOdfNodes(chart).find((node) => node.name === "table:table");
  const rows = (table?.children ?? []).filter(
    (node): node is OdfXmlNode => typeof node !== "string",
  );
  const firstColumn = columnNumber(match[1]!);
  const firstRow = Number(match[2]);
  const lastColumn = match[3] ? columnNumber(match[3]) : firstColumn;
  const lastRow = match[4] ? Number(match[4]) : firstRow;
  const values: string[] = [];
  for (let row = firstRow; row <= lastRow; row += 1) {
    const cells = (rows[row - 1]?.children ?? []).filter(
      (node): node is OdfXmlNode => typeof node !== "string",
    );
    for (let column = firstColumn; column <= lastColumn; column += 1) {
      const cell = cells[column - 1];
      values.push(typeof cell === "object" ? String(cell.attributes?.["office:value"] ?? "") : "");
    }
  }
  return values;
}

function columnNumber(column: string): number {
  let total = 0;
  for (let index = 0; index < column.length; index += 1) {
    total = total * 26 + column.charCodeAt(index) - 64;
  }
  return total;
}
