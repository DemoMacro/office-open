import type {
  AxisOptions,
  ChartSeriesData,
  ChartSpaceOptions,
  ChartTitleOptions,
  ScatterSeriesData,
  ShapePropertiesOptions,
} from "@office-open/core";
import { parse, type Element } from "@office-open/xml";

import type { OcfManifestOptions } from "../container/manifest";
import { generateOcf, readOcf, readXml, type OdfFiles } from "../container/package";
import {
  graphicFill,
  graphicOutline,
  hexColorValue,
  parseGraphicStyles,
  pushShapeStyle,
  type GraphicStyle,
} from "../drawing/graphic-style";
import { OdfSchemaError } from "../errors";
import { metaXml, ODF_NAMESPACES, parseMeta } from "../meta/core-properties";
import { parseOdfNode, parseOdfNodes, type OdfXmlNode } from "../runtime/odf-node";
import {
  attributeString,
  childNamed,
  childrenNamed,
  escapeText,
  textOf,
  xmlElement,
} from "../runtime/xml";

const MIME = "application/vnd.oasis.opendocument.chart";
export const CHART_MIME = MIME;
const NAMESPACES = `${ODF_NAMESPACES} xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"`;
const CONTENT_PATH = "content.xml";

/** Standalone ODF chart subdocument using the shared chart model. */
export interface ChartDocumentOptions {
  title?: string;
  chart: ChartSpaceOptions;
}

interface ChartStyle {
  graphic?: GraphicStyle;
  styleName?: string;
  properties: Record<string, string>;
}

interface ChartSeriesSource {
  element: Element;
  classToken: string;
  vertical: boolean;
  valuesRange?: string;
  labelRange?: string;
  attachedAxis?: string;
  styleName?: string;
  domains: string[];
  repeatedPoints: Array<{ repeated: number; styleName?: string }>;
}

const CHART_CLASSES: Record<ChartSpaceOptions["type"], string> = {
  column: "chart:bar",
  bar: "chart:bar",
  line: "chart:line",
  pie: "chart:circle",
  area: "chart:area",
  scatter: "chart:scatter",
  bubble: "chart:bubble",
  doughnut: "chart:ring",
  radar: "chart:radar",
  stock: "chart:stock",
  surface: "chart:fill",
  ofPie: "chart:circle",
};

const CLASS_CHARTS: Record<string, ChartSpaceOptions["type"]> = {
  "chart:bar": "bar",
  "chart:line": "line",
  "chart:circle": "pie",
  "chart:ring": "doughnut",
  "chart:area": "area",
  "chart:scatter": "scatter",
  "chart:bubble": "bubble",
  "chart:radar": "radar",
  "chart:filled-radar": "radar",
  "chart:stock": "stock",
  "chart:fill": "surface",
};

const LEGEND_POSITIONS: Record<string, ChartSpaceOptions["legendPosition"]> = {
  start: "left",
  end: "right",
  top: "top",
  bottom: "bottom",
  "top-end": "topRight",
};

export function generateChartDocument(options: ChartDocumentOptions): Uint8Array {
  return generateOcf(MIME, {
    "content.xml": chartBodyXml(options.chart),
    "meta.xml": metaXml({ title: options.title }),
  });
}

/** Serializes an embedded chart subdocument body (its content.xml). */
export function chartBodyXml(chart: ChartSpaceOptions): string {
  const styles: string[] = [];
  const body = chartXml(chart, styles);
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.3"><office:automatic-styles>${styles.join(
    "",
  )}</office:automatic-styles><office:body><office:chart>${body}</office:chart></office:body></office:document-content>`;
}

export function parseChartDocument(data: Uint8Array): ChartDocumentOptions {
  const { files } = readOcf(data, MIME);
  const content = readXml(files, CONTENT_PATH);
  const body = childNamed(childNamed(content, "office:body"), "office:chart");
  const chart = parseChart(
    childNamed(body, "chart:chart"),
    childNamed(content, "office:automatic-styles"),
  );
  if (!chart) {
    throw unsupported(
      CONTENT_PATH,
      "/office:document-content/office:body/office:chart",
      "office:chart",
      "missing chart root",
    );
  }
  return { ...parseMeta(files), chart };
}

/** Parses an embedded chart body; undefined when the XML is not a chart document. */
export function parseChartBody(xml: string): ChartSpaceOptions | undefined {
  const document = parse(xml, { ignoreDeclaration: true, ignoreDoctype: true });
  const root = document.elements?.[0] ?? document;
  const body = childNamed(childNamed(root, "office:body"), "office:chart");
  return parseChart(childNamed(body, "chart:chart"), childNamed(root, "office:automatic-styles"));
}

/** Embedded chart subdocuments from the manifest, keyed by their object name. */
export function parseEmbeddedCharts(
  manifest: OcfManifestOptions,
  files: OdfFiles,
): Map<string, ChartSpaceOptions> {
  const result = new Map<string, ChartSpaceOptions>();
  for (const entry of manifest.entries) {
    const fullPath = entry.fullPath;
    if (fullPath === "/" || !fullPath.endsWith("/")) continue;
    const content = files[`${fullPath}content.xml`];
    if (!content) {
      if (entry.mediaType === MIME) {
        throw unsupported(
          "META-INF/manifest.xml",
          `/manifest:manifest/manifest:file-entry[@manifest:full-path="${fullPath}"]`,
          "manifest:file-entry",
          "chart subdocument is missing content.xml",
        );
      }
      continue;
    }
    const chart = parseChartBody(content);
    if (entry.mediaType === MIME && !chart) {
      throw unsupported(
        `${fullPath}content.xml`,
        "/office:document-content/office:body/office:chart",
        "office:chart",
        "referenced chart subdocument is missing its chart body",
      );
    }
    if (chart && entry.mediaType !== undefined && entry.mediaType !== MIME) {
      throw unsupported(
        "META-INF/manifest.xml",
        `/manifest:manifest/manifest:file-entry[@manifest:full-path="${fullPath}"]`,
        "manifest:media-type",
        "chart subdocument manifest entry has an incompatible media type",
      );
    }
    if (!chart) {
      continue;
    }
    result.set(fullPath.replace(/\/$/, ""), chart);
  }
  for (const [path, content] of Object.entries(files)) {
    const match = /^(.+)\/content\.xml$/.exec(path);
    if (!match?.[1] || !content.includes("<office:chart")) continue;
    const objectPath = match[1];
    if (
      !result.has(objectPath) &&
      !manifest.entries.some((entry) => entry.fullPath === `${objectPath}/`)
    ) {
      throw unsupported(
        "META-INF/manifest.xml",
        `/manifest:manifest/manifest:file-entry[@manifest:full-path="${objectPath}/"]`,
        "manifest:file-entry",
        "chart subdocument path is not declared",
      );
    }
  }
  return result;
}

function chartXml(chart: ChartSpaceOptions, styles: string[]): string {
  const table = localTable(chart);
  return xmlElement(
    "chart:chart",
    {
      "chart:class": CHART_CLASSES[chart.type],
      "chart:style-name": pushShapeStyle(
        chart.shapeProperties?.fill,
        chart.shapeProperties?.outline,
        styles,
        "chart",
      ),
    },
    [
      titleXml("chart:title", chart.title, styles),
      legendXml(chart, styles),
      plotAreaXml(chart, table, styles),
      ...(table ? [serializeTable(table)] : []),
    ],
  );
}

function serializeTable(table: OdfXmlNode): string {
  const children = (table.children ?? []).map((child) =>
    typeof child === "string" ? escapeText(child) : serializeTable(child),
  );
  return xmlElement(table.name, table.attributes, children);
}

function titleText(title: string | ChartTitleOptions | undefined): string | undefined {
  if (typeof title === "string") return title;
  return typeof title?.text === "string" ? title.text : undefined;
}

function titleXml(
  name: string,
  title: string | ChartTitleOptions | undefined,
  styles: string[],
): string {
  const text = titleText(title);
  if (text === undefined) return "";
  return xmlElement(
    name,
    {
      "chart:style-name":
        title && typeof title === "object"
          ? pushShapeStyle(
              title.shapeProperties?.fill,
              title.shapeProperties?.outline,
              styles,
              "chart",
            )
          : undefined,
    },
    [xmlElement("text:p", undefined, [escapeText(text)])],
  );
}

function legendXml(chart: ChartSpaceOptions, styles: string[]): string {
  if (chart.showLegend === false) return "";
  const position = chart.legendPosition ?? "right";
  const odfPosition = Object.entries(LEGEND_POSITIONS).find(([, value]) => value === position)?.[0];
  if (!odfPosition) {
    throw unsupported(
      CONTENT_PATH,
      "/chart:chart/chart:legend",
      "chart:legend",
      `legend position ${position} has no ODF mapping`,
    );
  }
  return xmlElement(
    "chart:legend",
    {
      "chart:legend-position": odfPosition,
      ...(chart.legendOverlay === true ? { "svg:x": "0cm", "svg:y": "0cm" } : {}),
      "chart:style-name": pushShapeStyle(
        chart.legendShapeProperties?.fill,
        chart.legendShapeProperties?.outline,
        styles,
        "chart",
      ),
    },
    [],
  );
}

function plotAreaXml(
  chart: ChartSpaceOptions,
  table: OdfXmlNode | undefined,
  styles: string[],
): string {
  return xmlElement(
    "chart:plot-area",
    {
      "chart:style-name": pushShapeStyle(
        chart.plotAreaShapeProperties?.fill,
        chart.plotAreaShapeProperties?.outline,
        styles,
        "chart",
      ),
    },
    [
      ...(chart.axes ?? defaultAxes()).map((axis) => axisXml(axis, chart, table, styles)),
      ...allSeries(chart).map((series, index) => seriesXml(chart, series, index, table, styles)),
      wallXml(chart.sideWall ?? chart.backWall, styles),
      wallXml(chart.floor, styles, "chart:floor"),
    ],
  );
}

function allSeries(chart: ChartSpaceOptions): ChartSpaceOptions["series"][number][] {
  return [
    ...(chart.series as readonly ChartSpaceOptions["series"][number][]),
    ...((chart.secondaryGroups ?? []).flatMap(
      (group) => group.series,
    ) as readonly ChartSpaceOptions["series"][number][]),
  ];
}

function axisXml(
  axis: AxisOptions,
  chart: ChartSpaceOptions,
  table: OdfXmlNode | undefined,
  styles: string[],
): string {
  const categoryRange =
    chart.categoryFormula ?? (table ? `local-table.A2:A${rowCount(table)}` : undefined);
  return xmlElement(
    "chart:axis",
    {
      "chart:dimension": axis.kind === "value" || axis.kind === "date" ? "y" : "x",
      "chart:name": titleText(axis.title),
      "chart:style-name": pushAxisStyle(axis, styles),
    },
    [
      titleXml("chart:title", axis.title, styles),
      axis.kind !== "value" && axis.kind !== "date" && categoryRange
        ? xmlElement("chart:categories", { "table:cell-range-address": categoryRange })
        : "",
      axis.majorGridlines ? '<chart:grid chart:class="major"/>' : "",
      axis.minorGridlines ? '<chart:grid chart:class="minor"/>' : "",
    ],
  );
}

function pushAxisStyle(axis: AxisOptions, styles: string[]): string | undefined {
  const graphicName = pushShapeStyle(
    axis.shapeProperties?.fill,
    axis.shapeProperties?.outline,
    styles,
    "chart",
  );
  const chartAttributes = {
    "chart:visible": axis.delete === true ? "false" : undefined,
    "chart:display-label": axis.tickLabelPosition === "none" ? "false" : undefined,
  };
  if (Object.values(chartAttributes).every((value) => value === undefined)) return graphicName;
  const name = `ch-axis-${styles.length + 1}`;
  const graphicXml = graphicName
    ? (styles
        .find((style) => style.includes(`style:name="${graphicName}"`))
        ?.replace(/^<style:style[^>]*>|<\/style:style>$/g, "") ?? "")
    : "";
  styles.push(
    xmlElement("style:style", { "style:name": name, "style:family": "chart" }, [
      xmlElement("style:chart-properties", chartAttributes),
      ...(graphicXml ? [graphicXml] : []),
    ]),
  );
  return name;
}

function wallXml(
  wall: { thickness?: number | string; shapeProperties?: ShapePropertiesOptions } | undefined,
  styles: string[],
  name = "chart:wall",
): string {
  if (!wall) return "";
  return xmlElement(name, {
    "chart:style-name": pushShapeStyle(
      wall.shapeProperties?.fill,
      wall.shapeProperties?.outline,
      styles,
      "chart",
    ),
  });
}

function seriesXml(
  chart: ChartSpaceOptions,
  series: ChartSpaceOptions["series"][number],
  seriesIndex: number,
  table: OdfXmlNode | undefined,
  styles: string[],
): string {
  const values =
    series.valueFormula ??
    ("bubbleSize" in series && table
      ? `local-table.${columnLetters(seriesIndex + 3)}2:${columnLetters(seriesIndex + 3)}${rowCount(table)}`
      : localValuesRange(seriesIndex, table));
  const label = series.nameFormula ?? localLabelRange(seriesIndex, table);
  return xmlElement(
    "chart:series",
    {
      "chart:values-cell-range-address": values,
      "chart:label-cell-address": label,
      "chart:class": seriesClass(chart, series),
      "chart:vertical": chart.type === "column" ? "true" : undefined,
      "chart:attached-axis": "primary-y",
      "chart:style-name": pushShapeStyle(
        series.shapeProperties?.fill,
        series.shapeProperties?.outline,
        styles,
        "chart",
      ),
    },
    [
      ...seriesDomains(series, seriesIndex, table).map((domain) =>
        xmlElement("chart:domain", { "table:cell-range-address": domain }),
      ),
      ...dataPointXml(series, styles),
    ],
  );
}

function seriesClass(
  chart: ChartSpaceOptions,
  series: ChartSpaceOptions["series"][number],
): string {
  const secondary = chart.secondaryGroups?.find((group) =>
    group.series.includes(series as ChartSeriesData),
  );
  return CHART_CLASSES[secondary?.type ?? chart.type];
}

function seriesDomains(
  series: ChartSpaceOptions["series"][number],
  seriesIndex: number,
  table: OdfXmlNode | undefined,
): string[] {
  if (!("xValues" in series) || !table) return [];
  const x = `local-table.A2:A${rowCount(table)}`;
  return "bubbleSize" in series ? [x, localValuesRange(seriesIndex, table)!] : [x];
}

function dataPointXml(series: ChartSpaceOptions["series"][number], styles: string[]): string[] {
  return dataPointRuns(series).map((run) =>
    xmlElement("chart:data-point", {
      "chart:repeated": run.repeated,
      "chart:style-name": pushShapeStyle(
        run.first.shapeProperties?.fill,
        run.first.shapeProperties?.outline,
        styles,
        "chart",
      ),
    }),
  );
}

function dataPointRuns(
  series: ChartSpaceOptions["series"][number],
): Array<{ repeated: number; first: { shapeProperties?: ShapePropertiesOptions } }> {
  const runs: Array<{ repeated: number; first: { shapeProperties?: ShapePropertiesOptions } }> = [];
  for (const point of series.dataPoints ?? []) {
    const last = runs.at(-1);
    if (last && sameShape(last.first.shapeProperties, point.shapeProperties)) last.repeated += 1;
    else runs.push({ repeated: 1, first: point });
  }
  return runs;
}

function sameShape(
  left: ShapePropertiesOptions | undefined,
  right: ShapePropertiesOptions | undefined,
): boolean {
  return (
    hexColorValue(left?.fill) === hexColorValue(right?.fill) &&
    hexColorValue(left?.outline?.color) === hexColorValue(right?.outline?.color)
  );
}

function defaultAxes(): AxisOptions[] {
  return [{ kind: "category" }, { kind: "value", majorGridlines: true }];
}

function localValuesRange(seriesIndex: number, table: OdfXmlNode | undefined): string | undefined {
  if (!table) return undefined;
  const letter = columnLetters(seriesIndex + 2);
  return letter ? `local-table.${letter}2:${letter}${rowCount(table)}` : undefined;
}

function localLabelRange(seriesIndex: number, table: OdfXmlNode | undefined): string | undefined {
  if (!table) return undefined;
  const letter = columnLetters(seriesIndex + 2);
  return letter ? `local-table.${letter}1` : undefined;
}

function localTable(chart: ChartSpaceOptions): OdfXmlNode | undefined {
  if (chart.series.some((series) => series.valueFormula)) return undefined;
  const scatter = chart.type === "scatter" || chart.type === "bubble";
  const firstSeries = chart.series[0];
  const rows = Math.max(
    0,
    ...chart.series.map((series) =>
      "xValues" in series ? series.xValues.length : seriesValues(series).length,
    ),
  );
  if (rows === 0) return undefined;
  const header = scatter
    ? [
        "X",
        ...chart.series.map((series, index) => series.name ?? `Y ${index + 1}`),
        ...(chart.type === "bubble" ? ["Size"] : []),
      ]
    : ["", ...chart.series.map((series, index) => series.name ?? `Series ${index + 1}`)];
  const children = [
    tableRow(header),
    ...Array.from({ length: rows }, (_, row) =>
      tableRow(
        scatter
          ? [
              String(
                chart.series[0] && "xValues" in chart.series[0]
                  ? (chart.series[0].xValues[row] ?? "")
                  : "",
              ),
              ...chart.series.map((series) =>
                String(
                  ("yValues" in series ? series.yValues[row] : seriesValues(series)[row]) ?? "",
                ),
              ),
              ...(chart.type === "bubble"
                ? [
                    String(
                      firstSeries !== undefined && "bubbleSize" in firstSeries
                        ? (firstSeries.bubbleSize[row] ?? "")
                        : "",
                    ),
                  ]
                : []),
            ]
          : [
              chart.categories?.[row] ?? "",
              ...chart.series.map((series) => String(seriesValues(series)[row] ?? "")),
            ],
      ),
    ),
  ];
  return { name: "table:table", attributes: { "table:name": "local-table" }, children };
}

function tableRow(values: string[]): OdfXmlNode {
  return {
    name: "table:table-row",
    children: values.map((value) => ({
      name: "table:table-cell",
      attributes: {
        "office:value-type": value !== "" && Number.isFinite(Number(value)) ? "float" : "string",
        ...(value !== "" && Number.isFinite(Number(value)) ? { "office:value": value } : {}),
      },
      children: [value],
    })),
  };
}

function parseChart(
  element: Element | undefined,
  styleContainer: Element | undefined,
): ChartSpaceOptions | undefined {
  if (!element) return undefined;
  const classToken = attributeString(element, "chart:class");
  const type = classToken ? CLASS_CHARTS[classToken] : undefined;
  if (!type)
    throw unsupported(
      CONTENT_PATH,
      chartPath(element),
      element.name ?? "",
      `unsupported chart class ${classToken ?? "(missing)"}`,
    );
  const styleMap = chartStyles(styleContainer);
  const plotArea = childNamed(element, "chart:plot-area");
  if (!plotArea)
    throw unsupported(
      CONTENT_PATH,
      chartPath(element, "chart:plot-area"),
      "chart:plot-area",
      "missing plot area",
    );
  rejectUnknown(plotArea, ["chart:axis", "chart:series", "chart:wall", "chart:floor"]);
  rejectUnknown(element, [
    "chart:title",
    "chart:subtitle",
    "chart:footer",
    "chart:legend",
    "chart:plot-area",
    "table:table",
  ]);
  for (const name of ["chart:subtitle", "chart:footer"]) {
    if (childNamed(element, name))
      throw unsupported(
        CONTENT_PATH,
        chartPath(element, name),
        name,
        "the core chart model has no subtitle or footer field",
      );
  }
  const sources = childrenNamed(plotArea, "chart:series").map(parseSeriesSource);
  const series = sources.map((source, index) =>
    parseSeries(source, index, element, styleMap, sources, type),
  );
  const axes = childrenNamed(plotArea, "chart:axis").map((axis) => parseAxis(axis, styleMap));
  const categories = parseCategories(plotArea, element);
  return {
    type: sourceType(sources, type),
    title: parseTitle(childNamed(element, "chart:title"), styleMap),
    ...(categories ? { categories, categoryFormula: categoryRange(plotArea) } : {}),
    series: series as ChartSpaceOptions["series"],
    showLegend: childNamed(element, "chart:legend") !== undefined,
    legendPosition: parseLegendPosition(childNamed(element, "chart:legend")),
    legendOverlay: legendOverlay(childNamed(element, "chart:legend")),
    legendShapeProperties: styleShape(
      styleMap.get(attributeString(childNamed(element, "chart:legend"), "chart:style-name") ?? ""),
    ),
    shapeProperties: styleShape(styleMap.get(attributeString(element, "chart:style-name") ?? "")),
    axes,
    plotAreaShapeProperties: styleShape(
      styleMap.get(attributeString(plotArea, "chart:style-name") ?? ""),
    ),
    sideWall: parseWall(childNamed(plotArea, "chart:wall"), styleMap),
    floor: parseWall(childNamed(plotArea, "chart:floor"), styleMap),
  };
}

function sourceType(
  sources: ChartSeriesSource[],
  fallback: ChartSpaceOptions["type"],
): ChartSpaceOptions["type"] {
  const first = sources[0];
  if (!first) return fallback;
  if (first.classToken === "chart:bar") return first.vertical ? "column" : "bar";
  return CLASS_CHARTS[first.classToken] ?? fallback;
}

function parseSeriesSource(element: Element): ChartSeriesSource {
  rejectUnknown(element, [
    "chart:domain",
    "chart:data-point",
    "chart:mean-value",
    "chart:regression-curve",
    "chart:error-indicator",
    "chart:data-label",
  ]);
  for (const name of [
    "chart:mean-value",
    "chart:regression-curve",
    "chart:error-indicator",
    "chart:data-label",
  ]) {
    if (childNamed(element, name))
      throw unsupported(
        CONTENT_PATH,
        chartPath(element, name),
        name,
        "this series decoration has no canonical mapping in the source chart",
      );
  }
  return {
    element,
    classToken: attributeString(element, "chart:class") ?? "",
    vertical: attributeString(element, "chart:vertical") === "true",
    valuesRange: attributeString(element, "chart:values-cell-range-address"),
    labelRange: attributeString(element, "chart:label-cell-address"),
    attachedAxis: attributeString(element, "chart:attached-axis"),
    styleName: attributeString(element, "chart:style-name"),
    domains: childrenNamed(element, "chart:domain")
      .map((domain) => attributeString(domain, "table:cell-range-address") ?? "")
      .filter(Boolean),
    repeatedPoints: childrenNamed(element, "chart:data-point").map((point) => ({
      repeated: Number(attributeString(point, "chart:repeated") ?? "1"),
      styleName: attributeString(point, "chart:style-name"),
    })),
  };
}

function parseSeries(
  source: ChartSeriesSource,
  index: number,
  chart: Element,
  styleMap: Map<string, ChartStyle>,
  sources: ChartSeriesSource[],
  chartType: ChartSpaceOptions["type"],
): ChartSpaceOptions["series"][number] {
  const table = parseOdfNodes(chart).find((node) => node.name === "table:table");
  const type =
    index === 0 ? sourceType(sources, chartType) : (CLASS_CHARTS[source.classToken] ?? chartType);
  const scatter = type === "scatter" || type === "bubble";
  const values = localColumnValues(source.valuesRange, table);
  if (scatter) {
    const xValues = localColumnValues(source.domains[0], table);
    const yValues = localColumnValues(
      scatter ? (source.domains[1] ?? source.valuesRange) : source.valuesRange,
      table,
    );
    if (
      source.domains.some((domain) => !domain.startsWith("local-table.")) &&
      (xValues.length === 0 || yValues.length === 0)
    ) {
      throw unsupported(
        CONTENT_PATH,
        chartPath(source.element, "chart:domain"),
        "chart:domain",
        "external scatter X/Y formulas cannot be represented without cached values",
      );
    }
    const result: ScatterSeriesData = {
      name: source.labelRange ? localCell(chart, source.labelRange) : undefined,
      nameFormula: source.labelRange?.startsWith("local-table.") ? undefined : source.labelRange,
      valueFormula: source.valuesRange?.startsWith("local-table.") ? undefined : source.valuesRange,
      xValues,
      yValues,
      shapeProperties: styleShape(styleMap.get(source.styleName ?? "")),
    };
    if (type === "bubble") {
      return { ...result, bubbleSize: values, bubble3D: false };
    }
    return result;
  }
  return {
    name: source.labelRange ? localCell(chart, source.labelRange) : undefined,
    nameFormula: source.labelRange?.startsWith("local-table.") ? undefined : source.labelRange,
    valueFormula: source.valuesRange?.startsWith("local-table.") ? undefined : source.valuesRange,
    values,
    shapeProperties: styleShape(styleMap.get(source.styleName ?? "")),
    dataPoints: expandDataPoints(source, styleMap),
  };
}

function expandDataPoints(source: ChartSeriesSource, styleMap: Map<string, ChartStyle>) {
  let index = 0;
  return source.repeatedPoints.flatMap((run) =>
    Array.from({ length: Math.max(1, run.repeated) }, () => ({
      index: index++,
      shapeProperties: styleShape(styleMap.get(run.styleName ?? "")),
    })),
  );
}

function parseAxis(element: Element, styleMap: Map<string, ChartStyle>): AxisOptions {
  rejectUnknown(element, ["chart:title", "chart:categories", "chart:grid"]);
  const style = styleMap.get(attributeString(element, "chart:style-name") ?? "");
  const title =
    parseTitle(childNamed(element, "chart:title"), styleMap) ??
    attributeString(element, "chart:name");
  return {
    kind: attributeString(element, "chart:dimension") === "y" ? "value" : "category",
    title,
    delete: style?.properties["chart:visible"] === "false",
    tickLabelPosition: style?.properties["chart:display-label"] === "false" ? "none" : undefined,
    majorGridlines: childrenNamed(element, "chart:grid").some(
      (grid) => attributeString(grid, "chart:class") !== "minor",
    ),
    minorGridlines: childrenNamed(element, "chart:grid").some(
      (grid) => attributeString(grid, "chart:class") === "minor",
    ),
    shapeProperties: styleShape(style),
  };
}

function parseTitle(
  element: Element | undefined,
  styleMap: Map<string, ChartStyle>,
): string | ChartTitleOptions | undefined {
  if (!element) return undefined;
  const text = textOf(childNamed(element, "text:p"));
  if (!text) return undefined;
  const shapeProperties = styleShape(
    styleMap.get(attributeString(element, "chart:style-name") ?? ""),
  );
  return shapeProperties ? { text, shapeProperties } : text;
}

function parseLegendPosition(element: Element | undefined): ChartSpaceOptions["legendPosition"] {
  const position = attributeString(element, "chart:legend-position");
  const mapped = position ? LEGEND_POSITIONS[position] : undefined;
  if (element && position && !mapped)
    throw unsupported(
      CONTENT_PATH,
      chartPath(element),
      element.name ?? "",
      `legend position ${position} has no canonical mapping`,
    );
  return mapped ?? "right";
}

function legendOverlay(element: Element | undefined): boolean | undefined {
  if (!element) return undefined;
  return attributeString(element, "svg:x") !== undefined ||
    attributeString(element, "svg:y") !== undefined
    ? true
    : false;
}

function parseWall(element: Element | undefined, styleMap: Map<string, ChartStyle>) {
  if (!element) return undefined;
  return {
    shapeProperties: styleShape(styleMap.get(attributeString(element, "chart:style-name") ?? "")),
  };
}

function parseCategories(plotArea: Element | undefined, chart: Element): string[] | undefined {
  const axis = childrenNamed(plotArea, "chart:axis").find(
    (item) => attributeString(item, "chart:dimension") !== "y",
  );
  const range = attributeString(childNamed(axis, "chart:categories"), "table:cell-range-address");
  if (!range) return undefined;
  return localCells(chart, range).filter((value) => value !== "");
}

function categoryRange(plotArea: Element | undefined): string | undefined {
  const axis = childrenNamed(plotArea, "chart:axis").find(
    (item) => attributeString(item, "chart:dimension") !== "y",
  );
  const range = attributeString(childNamed(axis, "chart:categories"), "table:cell-range-address");
  return range?.startsWith("local-table.") ? undefined : range;
}

function chartStyles(container: Element | undefined): Map<string, ChartStyle> {
  const graphics = parseGraphicStyles(container, "chart");
  const result = new Map<string, ChartStyle>();
  for (const style of childrenNamed(container, "style:style")) {
    if (attributeString(style, "style:family") !== "chart") continue;
    const name = attributeString(style, "style:name") ?? "";
    const properties = childNamed(style, "style:chart-properties");
    result.set(name, {
      graphic: graphics.get(name),
      styleName: name,
      properties: Object.fromEntries(
        Object.entries(properties?.attributes ?? {}).map(([key, value]) => [key, String(value)]),
      ),
    });
  }
  return result;
}

function styleShape(style: ChartStyle | undefined): ShapePropertiesOptions | undefined {
  if (!style?.graphic) return undefined;
  const unsupportedFill =
    style.graphic.fill !== undefined &&
    style.graphic.fill !== "none" &&
    style.graphic.fill !== "solid" &&
    !style.graphic.fillColor;
  const unsupportedStroke =
    style.graphic.stroke !== undefined &&
    style.graphic.stroke !== "none" &&
    style.graphic.stroke !== "solid";
  if (unsupportedFill || unsupportedStroke) {
    const property = unsupportedFill ? "draw:fill" : "draw:stroke";
    const value = unsupportedFill ? style.graphic.fill : style.graphic.stroke;
    throw unsupported(
      CONTENT_PATH,
      `/office:document-content/office:automatic-styles/style:style[@style:name="${style.styleName ?? ""}"]/style:graphic-properties`,
      property,
      `${value} has no canonical ChartSpaceOptions mapping`,
    );
  }
  const fill = graphicFill(style.graphic);
  const outline = graphicOutline(style.graphic);
  return fill || outline
    ? { ...(fill ? { fill } : {}), ...(outline ? { outline } : {}) }
    : undefined;
}

function localColumnValues(range: string | undefined, table: OdfXmlNode | undefined): number[] {
  if (!range || !table) return [];
  const element = parse(serializeTable(table), { ignoreDeclaration: true });
  return localCells(element, range).map(Number).filter(Number.isFinite);
}

function localCell(chart: Element, range: string): string | undefined {
  return localCells(chart, range)[0];
}

function localCells(chart: Element, range: string | undefined): string[] {
  if (!range) return [];
  const match = /^local-table\.([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(range);
  if (!match) return [];
  const table =
    chart.name === "table:table"
      ? parseOdfNode(chart)
      : parseOdfNodes(chart).find((node) => node.name === "table:table");
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
      values.push(
        typeof cell === "object"
          ? String(
              cell.attributes?.["office:value"] ??
                (cell.children ?? []).find((child) => typeof child === "string") ??
                "",
            )
          : "",
      );
    }
  }
  return values;
}

function seriesValues(series: ChartSpaceOptions["series"][number]): readonly number[] {
  return "values" in series ? series.values : series.yValues;
}

function rowCount(table: OdfXmlNode): number {
  return (table.children ?? []).filter((child) => typeof child !== "string").length;
}

function columnLetters(number: number): string {
  let value = number;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function columnNumber(column: string): number {
  let total = 0;
  for (let index = 0; index < column.length; index += 1)
    total = total * 26 + column.charCodeAt(index) - 64;
  return total;
}

function chartPath(element: Element, child?: string): string {
  const base = "/office:document-content/office:body/office:chart/chart:chart";
  return child ? `${base}/${child}` : `${base}[${element.name}]`;
}

function rejectUnknown(element: Element, allowed: string[]): void {
  for (const child of element.elements ?? []) {
    if (child.type === "text" || child.type === "comment" || child.name === undefined) continue;
    if (!allowed.includes(child.name)) {
      throw unsupported(
        CONTENT_PATH,
        `${chartPath(element)}/${child.name}`,
        child.name,
        "element has no canonical ChartSpaceOptions mapping",
      );
    }
  }
}

function unsupported(part: string, path: string, name: string, reason: string): OdfSchemaError {
  return new OdfSchemaError(`${part}: ${path}: ${name}: ${reason}`, part, path, name, reason);
}
