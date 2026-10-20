import type {
  AxisOptions,
  ChartCoordinateRegionOptions,
  ChartLinesOptions,
  ChartShapePropertiesOptions,
  ChartSeriesData,
  ChartSpaceOptions,
  ChartTitleOptions,
  ChartValue,
  OdfDr3dLightResource,
  DataLabelsOptions,
  ErrorBarOptions,
  DataType,
  DataTableOptions,
  ChartStyleResources,
  ScatterSeriesData,
  ShapePropertiesOptions,
  TrendlineOptions,
  TrendlineType,
  UpDownBarsOptions,
} from "@office-open/core";
import { parse, type Element } from "@office-open/xml";

import {
  chartMediaFiles,
  chartStylesXml,
  parseChartStyleContext,
  projectChartStyle,
  pushChartStyle,
  type ChartStyle,
} from "./chart-style";
import type { OcfManifestOptions } from "./container/manifest";
import {
  generateOcf,
  readOcf,
  readXml,
  type OdfFiles,
  type OdfPackageFiles,
} from "./container/package";
import { hexColorValue } from "./drawing/graphic-style";
import { OdfSchemaError } from "./errors";
import { metaXml, ODF_NAMESPACES, parseMeta } from "./meta";
import { parseOdfNode, parseOdfNodes, type OdfXmlNode } from "./runtime/odf-node";
import {
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  escapeText,
  lengthToEmu,
  spreadsheetColumnName as columnLetters,
  spreadsheetColumnNumber as columnNumber,
  textOf,
  xmlElement,
} from "./runtime/xml";

const MIME = "application/vnd.oasis.opendocument.chart";
export const CHART_MIME = MIME;
const NAMESPACES =
  `${ODF_NAMESPACES} xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"` +
  ' xmlns:chartooo="urn:org:documentfoundation:names:experimental:chart:xmlns:loext:1.0"' +
  ' xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"';
const CONTENT_PATH = "content.xml";

/** Standalone ODF chart subdocument using the shared chart model. */
export interface ChartDocumentOptions {
  title?: string;
  chart: ChartSpaceOptions;
}

interface ChartSeriesSource {
  element: Element;
  classToken: string;
  vertical: boolean;
  valuesRange?: string;
  labelRange?: string;
  attachedAxis?: string;
  styleName?: string;
  dataLabelStyleName?: string;
  domains: string[];
  repeatedPoints: Array<{ repeated: number; styleName?: string; labelStyleName?: string }>;
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
    ...chartSubdocumentFiles("", options.chart),
    "meta.xml": metaXml({ title: options.title }),
  });
}

/** Builds the content, style, and media parts for an embedded chart. */
export function chartSubdocumentFiles(basePath: string, chart: ChartSpaceOptions): OdfPackageFiles {
  const resources = chart.styleResources ?? emptyChartStyleResources();
  return {
    [`${basePath}content.xml`]: chartBodyXml(chart),
    [`${basePath}styles.xml`]: chartStylesXml(resources),
    ...chartMediaFiles(basePath, resources),
  };
}

/** Serializes an embedded chart subdocument body (its content.xml). */
export function chartBodyXml(chart: ChartSpaceOptions): string {
  const styles: string[] = [];
  const previousResources = activeChartStyleResources;
  activeChartStyleResources = chart.styleResources ?? emptyChartStyleResources();
  try {
    const body = chartXml(chart, styles);
    return `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.3"><office:automatic-styles>${styles.join(
      "",
    )}</office:automatic-styles><office:body><office:chart>${body}</office:chart></office:body></office:document-content>`;
  } finally {
    activeChartStyleResources = previousResources;
  }
}

let activeChartStyleResources = emptyChartStyleResources();

function pushChartShapeStyle(
  shapeProperties: ChartShapePropertiesOptions | undefined,
  styles: string[],
  fallbackName: string,
  chartProperties: Record<string, string> = {},
): string | undefined {
  return pushChartStyle(
    shapeProperties,
    activeChartStyleResources,
    styles,
    fallbackName,
    chartProperties,
  );
}

export async function parseChartDocument(data: DataType): Promise<ChartDocumentOptions> {
  const { files, binaries } = await readOcf(data, MIME);
  const content = readXml(files, CONTENT_PATH);
  const styles = files["styles.xml"] ? xmlRoot(files["styles.xml"]) : undefined;
  const context = parseChartStyleContext(
    [
      { part: CONTENT_PATH, root: content },
      ...(styles ? [{ part: "styles.xml", root: styles }] : []),
    ],
    binaries,
    "",
  );
  const body = childNamed(childNamed(content, "office:body"), "office:chart");
  const chart = parseChart(childNamed(body, "chart:chart"), context, "");
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
export function parseChartBody(xml: string, stylesXml?: string): ChartSpaceOptions | undefined {
  const root = xmlRoot(xml);
  const styles = stylesXml ? xmlRoot(stylesXml) : undefined;
  const context = parseChartStyleContext(
    [{ part: CONTENT_PATH, root }, ...(styles ? [{ part: "styles.xml", root: styles }] : [])],
    {},
    "",
  );
  const body = childNamed(childNamed(root, "office:body"), "office:chart");
  return parseChart(childNamed(body, "chart:chart"), context, "");
}

/** Embedded chart subdocuments from the manifest, keyed by their object name. */
export function parseEmbeddedCharts(
  manifest: OcfManifestOptions,
  files: OdfFiles,
  binaries: Record<string, Uint8Array> = {},
  packageStylesXml?: string,
): Map<string, ChartSpaceOptions> {
  const result = new Map<string, ChartSpaceOptions>();
  const packageStyles = packageStylesXml
    ? [{ part: "styles.xml", root: xmlRoot(packageStylesXml) }]
    : [];
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
    const styles = files[`${fullPath}styles.xml`];
    const contentRoot = xmlRoot(content);
    const stylesRoot = styles ? xmlRoot(styles) : undefined;
    const context = parseChartStyleContext(
      [
        { part: `${fullPath}content.xml`, root: contentRoot },
        ...(stylesRoot ? [{ part: `${fullPath}styles.xml`, root: stylesRoot }] : []),
        ...packageStyles,
      ],
      binaries,
      fullPath,
    );
    const body = childNamed(childNamed(contentRoot, "office:body"), "office:chart");
    const chart = parseChart(childNamed(body, "chart:chart"), context, fullPath);
    if (chart && entry.mediaType !== undefined && entry.mediaType !== MIME) {
      throw unsupported(
        "META-INF/manifest.xml",
        `/manifest:manifest/manifest:file-entry[@manifest:full-path="${fullPath}"]`,
        "manifest:media-type",
        "chart subdocument manifest entry has an incompatible media type",
      );
    }
    if (!chart) {
      if (entry.mediaType === MIME || body?.name === "office:chart") {
        throw unsupported(
          `${fullPath}content.xml`,
          "/office:document-content/office:body/office:chart",
          body?.name ?? "office:chart",
          "chart subdocument has no canonical chart root",
        );
      }
      continue;
    }
    result.set(fullPath.replace(/\/$/, ""), { ...chart, styleResources: context.resources });
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

/** Resolves a draw:object href against manifest keys, including URI escapes. */
export function resolveEmbeddedChart(
  charts: Map<string, ChartSpaceOptions>,
  href: string,
): ChartSpaceOptions | undefined {
  const normalizedHref = decodeURIComponent(
    href.replace(/^\.\//, "").replace(/^\//, "").replace(/\/$/, ""),
  );
  if (charts.has(normalizedHref)) return charts.get(normalizedHref);
  const hrefBase = normalizedHref.split("/").pop() ?? normalizedHref;
  for (const [path, chart] of charts) {
    const normalizedPath = decodeURIComponent(path.replace(/^\//, "").replace(/\/$/, ""));
    if (normalizedPath === normalizedHref || normalizedPath.endsWith(`/${hrefBase}`)) return chart;
  }
  return undefined;
}

function chartXml(chart: ChartSpaceOptions, styles: string[]): string {
  const table = localTable(chart);
  return xmlElement(
    "chart:chart",
    {
      "chart:class": CHART_CLASSES[chart.type],
      "chart:style-name": pushChartShapeStyle(chart.shapeProperties, styles, "chart"),
    },
    [
      titleXml("chart:title", chart.title, styles),
      titleXml("chart:subtitle", chart.subtitle, styles),
      titleXml("chart:footer", chart.footer, styles),
      legendXml(chart, styles),
      dataTableXml(chart.dataTable, styles),
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
          ? pushChartShapeStyle(title.shapeProperties, styles, "chart")
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
      "chart:style-name": pushChartShapeStyle(chart.legendShapeProperties, styles, "chart"),
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
      "chart:style-name": pushChartShapeStyle(chart.plotAreaShapeProperties, styles, "chart"),
    },
    [
      coordinateRegionXml(chart.coordinateRegion),
      ...(chart.axes ?? defaultAxes()).map((axis) => axisXml(axis, chart, table, styles)),
      stockBarsXml(typeof chart.upDownBars === "object" ? chart.upDownBars : undefined, styles),
      ...allSeries(chart).map((series, index) => seriesXml(chart, series, index, table, styles)),
      wallXml(chart.sideWall ?? chart.backWall, styles),
      wallXml(chart.floor, styles, "chart:floor"),
      ...(chart.styleResources?.dr3dLights ?? []).map(dr3dLightXml),
    ],
  );
}

function stockBarsXml(bars: UpDownBarsOptions | undefined, styles: string[]): string {
  return [
    bars?.upBars === false ? "" : stockBarXml("chart:stock-gain-marker", bars?.upBars, styles),
    bars?.downBars === false ? "" : stockBarXml("chart:stock-loss-marker", bars?.downBars, styles),
  ].join("");
}

function stockBarXml(
  name: string,
  line: boolean | ChartLinesOptions | undefined,
  styles: string[],
): string {
  return xmlElement(name, {
    "chart:style-name": pushChartShapeStyle(
      typeof line === "object" ? line.shapeProperties : undefined,
      styles,
      "stock",
    ),
  });
}

function dr3dLightXml(light: OdfDr3dLightResource): string {
  return xmlElement("dr3d:light", {
    "dr3d:diffuse-color": light.diffuseColor,
    "dr3d:direction": light.direction,
    "dr3d:enabled": light.enabled,
    "dr3d:specular": light.specular,
  });
}

function dataTableXml(table: DataTableOptions | undefined, styles: string[]): string {
  if (!table) return "";
  return xmlElement("loext:data-table", {
    "chart:style-name": pushChartShapeStyle(table.shapeProperties, styles, "table"),
  });
}

function coordinateRegionXml(region: ChartCoordinateRegionOptions | undefined): string {
  if (!region) return "";
  return xmlElement("chart:coordinate-region", {
    "svg:x": region.x === undefined ? undefined : emuToLength(region.x),
    "svg:y": region.y === undefined ? undefined : emuToLength(region.y),
    "svg:width": region.width === undefined ? undefined : emuToLength(region.width),
    "svg:height": region.height === undefined ? undefined : emuToLength(region.height),
  });
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
      "chart:dimension": axis.kind === "value" ? "y" : "x",
      "chart:name": titleText(axis.title),
      "chart:style-name": pushAxisStyle(axis, styles),
    },
    [
      titleXml("chart:title", axis.title, styles),
      axis.kind !== "value" && axis.kind !== "date" && categoryRange
        ? xmlElement("chart:categories", { "table:cell-range-address": categoryRange })
        : "",
      axis.kind === "date" ? "<chartooo:date-scale/>" : "",
      axis.majorGridlines ? '<chart:grid chart:class="major"/>' : "",
      axis.minorGridlines ? '<chart:grid chart:class="minor"/>' : "",
    ],
  );
}

function pushAxisStyle(axis: AxisOptions, styles: string[]): string | undefined {
  const graphicName = pushChartShapeStyle(axis.shapeProperties, styles, "chart");
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
    "chart:style-name": pushChartShapeStyle(wall.shapeProperties, styles, "chart"),
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
      "chart:attached-axis": "primary-y",
      "chart:style-name": pushChartShapeStyle(
        series.shapeProperties,
        styles,
        "series",
        chart.type === "column" ? { "chart:vertical": "true" } : {},
      ),
    },
    [
      ...seriesDomains(series, seriesIndex, table).map((domain) =>
        xmlElement("chart:domain", { "table:cell-range-address": domain }),
      ),
      ...dataPointXml(series, styles),
      ...(series.trendlines ?? []).map((trendline) =>
        xmlElement("chart:regression-curve", {
          "chart:style-name": pushChartShapeStyle(trendline.shapeProperties, styles, "chart"),
        }),
      ),
      ...(series.errorBars ?? []).map((errorBars) =>
        xmlElement("chart:error-indicator", {
          "chart:dimension": errorBars.direction ?? "y",
        }),
      ),
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
      "chart:style-name": pushChartShapeStyle(run.first.shapeProperties, styles, "chart"),
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

function seriesNameText(series: ChartSeriesData | ScatterSeriesData, index: number): string {
  if (typeof series.name === "string") return series.name;
  return series.name?.text ?? (series.name === undefined ? `Series ${index + 1}` : "");
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
        ...chart.series.map((series, index) => seriesNameText(series, index)),
        ...(chart.type === "bubble" ? ["Size"] : []),
      ]
    : ["", ...chart.series.map((series, index) => seriesNameText(series, index))];
  const children = [
    {
      name: "table:table-columns",
      children: [
        {
          name: "table:table-column",
          attributes: { "table:number-columns-repeated": Math.max(1, header.length) },
        },
      ],
    },
    tableRow(header),
    ...Array.from({ length: rows }, (_, row) =>
      tableRow(
        scatter
          ? [
              chartValueText(
                chart.series[0] && "xValues" in chart.series[0]
                  ? chart.series[0].xValues[row]
                  : undefined,
              ),
              ...chart.series.map((series) =>
                chartValueText(
                  "yValues" in series ? series.yValues[row] : seriesValues(series)[row],
                ),
              ),
              ...(chart.type === "bubble"
                ? [
                    chartValueText(
                      firstSeries !== undefined && "bubbleSize" in firstSeries
                        ? firstSeries.bubbleSize[row]
                        : undefined,
                    ),
                  ]
                : []),
            ]
          : [
              typeof chart.categories?.[row] === "object"
                ? (chart.categories[row]?.text ?? "")
                : (chart.categories?.[row] ?? ""),
              ...chart.series.map((series) => chartValueText(seriesValues(series)[row])),
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
      children: [{ name: "text:p", children: [value] }],
    })),
  };
}

function parseChart(
  element: Element | undefined,
  context: { resources: ChartStyleResources; styles: Map<string, ChartStyle> },
  basePath: string,
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
  const styleMap = context.styles;
  const plotArea = childNamed(element, "chart:plot-area");
  if (!plotArea)
    throw unsupported(
      CONTENT_PATH,
      chartPath(element, "chart:plot-area"),
      "chart:plot-area",
      "missing plot area",
    );
  rejectUnknown(plotArea, [
    "chart:coordinate-region",
    "chartooo:coordinate-region",
    "chart:axis",
    "chart:series",
    "chart:wall",
    "chart:floor",
    "chart:stock-gain-marker",
    "chart:stock-loss-marker",
    "dr3d:light",
  ]);
  rejectUnknown(element, [
    "chart:title",
    "chart:subtitle",
    "chart:footer",
    "chart:legend",
    "chart:plot-area",
    "table:table",
    "loext:data-table",
  ]);
  const sources = childrenNamed(plotArea, "chart:series").map(parseSeriesSource);
  for (const source of sources) {
    source.vertical = styleMap.get(source.styleName ?? "")?.properties["chart:vertical"] === "true";
  }
  const series = sources.map((source, index) =>
    parseSeries(source, index, element, styleMap, sources, type, context.resources, basePath),
  );
  for (const light of childrenNamed(plotArea, "dr3d:light"))
    (context.resources.dr3dLights ??= []).push({
      diffuseColor: attributeString(light, "dr3d:diffuse-color"),
      direction: attributeString(light, "dr3d:direction"),
      enabled: attributeString(light, "dr3d:enabled") === "true",
      specular: attributeString(light, "dr3d:specular") === "true",
    });
  const axes = childrenNamed(plotArea, "chart:axis").map((axis) =>
    parseAxis(axis, styleMap, context.resources, basePath),
  );
  const upDownBars = parseStockBars(plotArea, styleMap, context.resources, basePath);
  const dataTable = parseDataTable(
    childNamed(element, "loext:data-table"),
    styleMap,
    context.resources,
    basePath,
  );
  const categories = parseCategories(plotArea, element);
  return {
    type: sourceType(sources, type),
    ...(hasChartStyleResources(context.resources) ? { styleResources: context.resources } : {}),
    title: parseTitle(childNamed(element, "chart:title"), styleMap, context.resources, basePath),
    subtitle: parseTitle(
      childNamed(element, "chart:subtitle"),
      styleMap,
      context.resources,
      basePath,
    ),
    footer: parseTitle(childNamed(element, "chart:footer"), styleMap, context.resources, basePath),
    ...(categories ? { categories, categoryFormula: categoryRange(plotArea) } : {}),
    series: series as ChartSpaceOptions["series"],
    showLegend: childNamed(element, "chart:legend") !== undefined,
    legendPosition: parseLegendPosition(childNamed(element, "chart:legend")),
    legendOverlay: legendOverlay(childNamed(element, "chart:legend")),
    legendShapeProperties: chartShape(
      childNamed(element, "chart:legend"),
      styleMap,
      context.resources,
      basePath,
    ),
    shapeProperties: chartShape(element, styleMap, context.resources, basePath),
    axes,
    ...(upDownBars ? { upDownBars } : {}),
    ...(dataTable ? { dataTable } : {}),
    coordinateRegion: parseCoordinateRegion(
      childNamed(plotArea, "chart:coordinate-region") ??
        childNamed(plotArea, "chartooo:coordinate-region"),
    ),
    plotAreaShapeProperties: chartShape(plotArea, styleMap, context.resources, basePath),
    sideWall: parseWall(childNamed(plotArea, "chart:wall"), styleMap, context.resources, basePath),
    floor: parseWall(childNamed(plotArea, "chart:floor"), styleMap, context.resources, basePath),
  };
}

function parseStockBars(
  plotArea: Element,
  styleMap: Map<string, ChartStyle>,
  resources: ChartStyleResources,
  basePath: string,
): UpDownBarsOptions | undefined {
  const gain = childNamed(plotArea, "chart:stock-gain-marker");
  const loss = childNamed(plotArea, "chart:stock-loss-marker");
  if (!gain && !loss) return undefined;
  return {
    upBars: chartLine(gain, styleMap, resources, basePath),
    downBars: chartLine(loss, styleMap, resources, basePath),
  };
}

function parseDataTable(
  element: Element | undefined,
  styleMap: Map<string, ChartStyle>,
  resources: ChartStyleResources,
  basePath: string,
): DataTableOptions | undefined {
  if (!element) return undefined;
  return {
    shapeProperties: styleShape(
      requiredStyle(styleMap, attributeString(element, "chart:style-name"), element, basePath),
      resources,
      basePath,
    ),
  };
}

function chartLine(
  element: Element | undefined,
  styleMap: Map<string, ChartStyle>,
  resources: ChartStyleResources,
  basePath: string,
): boolean | ChartLinesOptions {
  if (!element) return false;
  const shapeProperties = styleShape(
    requiredStyle(styleMap, attributeString(element, "chart:style-name"), element, basePath),
    resources,
    basePath,
  );
  return shapeProperties ? { shapeProperties } : true;
}

function parseCoordinateRegion(
  element: Element | undefined,
): ChartCoordinateRegionOptions | undefined {
  if (!element) return undefined;
  const length = (name: string): number | string | undefined => {
    const value = attributeString(element, name);
    if (value === undefined) return undefined;
    const emu = lengthToEmu(value);
    if (emu === undefined) {
      throw unsupported(
        CONTENT_PATH,
        chartPath(element, name),
        element.name ?? "chart:coordinate-region",
        `${name} is not a valid ODF length`,
      );
    }
    return emu;
  };
  return {
    x: length("svg:x"),
    y: length("svg:y"),
    width: length("svg:width"),
    height: length("svg:height"),
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
  for (const name of ["chart:mean-value"]) {
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
    dataLabelStyleName: attributeString(
      childNamed(element, "chart:data-label"),
      "chart:style-name",
    ),
    domains: childrenNamed(element, "chart:domain")
      .map((domain) => attributeString(domain, "table:cell-range-address") ?? "")
      .filter(Boolean),
    repeatedPoints: childrenNamed(element, "chart:data-point").map((point) => ({
      repeated: Number(attributeString(point, "chart:repeated") ?? "1"),
      styleName: attributeString(point, "chart:style-name"),
      labelStyleName: attributeString(childNamed(point, "chart:data-label"), "chart:style-name"),
    })),
  };
}

function requiredStyle(
  styleMap: Map<string, ChartStyle>,
  styleName: string | undefined,
  element: Element,
  basePath: string,
): ChartStyle | undefined {
  if (!styleName) return undefined;
  const style = styleMap.get(styleName);
  if (!style) {
    throw unsupported(
      `${basePath}content.xml`,
      `${chartPath(element)}/@chart:style-name`,
      styleName,
      "unresolved-style-reference",
    );
  }
  return style;
}

function seriesShape(
  source: ChartSeriesSource,
  styleMap: Map<string, ChartStyle>,
  resources: ChartStyleResources,
  basePath: string,
): ChartShapePropertiesOptions | undefined {
  return styleShape(
    requiredStyle(styleMap, source.styleName, source.element, basePath),
    resources,
    basePath,
  );
}

function seriesDataLabels(
  source: ChartSeriesSource,
  styleMap: Map<string, ChartStyle>,
  resources: ChartStyleResources,
  basePath: string,
): DataLabelsOptions | undefined {
  const seriesLabel = source.dataLabelStyleName;
  let index = 0;
  const labels = source.repeatedPoints.flatMap((point) =>
    Array.from({ length: Math.max(1, point.repeated) }, () => {
      const current = index++;
      if (!point.labelStyleName) return undefined;
      return {
        index: current,
        shapeProperties: styleShape(
          requiredStyle(styleMap, point.labelStyleName, source.element, basePath),
          resources,
          basePath,
        ),
      };
    }),
  );
  const shapeProperties = seriesLabel
    ? styleShape(
        requiredStyle(styleMap, seriesLabel, source.element, basePath),
        resources,
        basePath,
      )
    : undefined;
  if (!seriesLabel && labels.length === 0) return undefined;
  const pointLabels = labels.filter(
    (label): label is NonNullable<typeof label> => label !== undefined,
  );
  return {
    ...(shapeProperties ? { shapeProperties } : {}),
    ...(pointLabels.length > 0 ? { labels: pointLabels } : {}),
  };
}

function parseSeries(
  source: ChartSeriesSource,
  index: number,
  chart: Element,
  styleMap: Map<string, ChartStyle>,
  sources: ChartSeriesSource[],
  chartType: ChartSpaceOptions["type"],
  resources: ChartStyleResources,
  basePath: string,
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
    const result: ScatterSeriesData = {
      name: source.labelRange ? localCell(chart, source.labelRange) : undefined,
      nameFormula: source.labelRange?.startsWith("local-table.") ? undefined : source.labelRange,
      valueFormula: source.valuesRange?.startsWith("local-table.") ? undefined : source.valuesRange,
      xValues,
      yValues,
      shapeProperties: seriesShape(source, styleMap, resources, basePath),
      trendlines: parseTrendlines(styleMap.get(source.styleName ?? "")),
      errorBars: parseErrorBars(source.element, styleMap.get(source.styleName ?? "")),
      dataLabels: seriesDataLabels(source, styleMap, resources, basePath),
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
    shapeProperties: seriesShape(source, styleMap, resources, basePath),
    trendlines: parseTrendlines(styleMap.get(source.styleName ?? "")),
    errorBars: parseErrorBars(source.element, styleMap.get(source.styleName ?? "")),
    dataLabels: seriesDataLabels(source, styleMap, resources, basePath),
    dataPoints: expandDataPoints(source, styleMap, resources, basePath),
  };
}

function expandDataPoints(
  source: ChartSeriesSource,
  styleMap: Map<string, ChartStyle>,
  resources: ChartStyleResources,
  basePath: string,
) {
  let index = 0;
  return source.repeatedPoints.flatMap((run) =>
    Array.from({ length: Math.max(1, run.repeated) }, () => ({
      index: index++,
      shapeProperties: styleShape(
        requiredStyle(styleMap, run.styleName, source.element, basePath),
        resources,
        basePath,
      ),
    })),
  );
}

function parseTrendlines(style: ChartStyle | undefined): TrendlineOptions[] | undefined {
  const type = style?.properties["chart:regression-type"];
  if (!type || type === "none") return undefined;
  const trendline: TrendlineOptions = {
    type:
      type === "moving-average"
        ? "movingAverage"
        : (type as Exclude<TrendlineType, "movingAverage">),
  };
  const degree = style?.properties["chart:regression-max-degree"];
  if (type === "polynomial" && degree !== undefined) trendline.order = Number(degree);
  if (style?.properties["chart:regression-force-intercept"] === "true") {
    trendline.intercept = Number(style?.properties["chart:regression-intercept-value"] ?? 0);
  }
  return [trendline];
}

function parseErrorBars(
  element: Element,
  style: ChartStyle | undefined,
): ErrorBarOptions[] | undefined {
  const indicators = childrenNamed(element, "chart:error-indicator");
  if (indicators.length === 0) return undefined;
  return indicators.map((indicator) => {
    const category = style?.properties["chart:error-category"] ?? "none";
    const upper = style?.properties["chart:error-upper-indicator"] !== "false";
    const lower = style?.properties["chart:error-lower-indicator"] !== "false";
    const result: ErrorBarOptions = {
      direction: attributeString(indicator, "chart:dimension") === "x" ? "x" : "y",
      barType: upper && lower ? "both" : upper ? "plus" : "minus",
      valueType:
        category === "standard-deviation"
          ? "standardDeviation"
          : category === "percentage"
            ? "percentage"
            : category === "standard-error"
              ? "standardError"
              : category === "none"
                ? "custom"
                : "fixedValue",
    };
    if (result.valueType === "fixedValue") {
      result.value = Number(
        style?.properties["chart:error-margin"] ??
          style?.properties["chart:error-lower-limit"] ??
          0,
      );
    }
    return result;
  });
}

function parseAxis(
  element: Element,
  styleMap: Map<string, ChartStyle>,
  resources: ChartStyleResources,
  basePath: string,
): AxisOptions {
  rejectUnknown(element, ["chart:title", "chart:categories", "chart:grid", "chartooo:date-scale"]);
  const style = requiredStyle(
    styleMap,
    attributeString(element, "chart:style-name"),
    element,
    basePath,
  );
  const title =
    parseTitle(childNamed(element, "chart:title"), styleMap, resources, basePath) ??
    attributeString(element, "chart:name");
  return {
    kind:
      attributeString(element, "chart:dimension") === "y"
        ? "value"
        : childNamed(element, "chartooo:date-scale")
          ? "date"
          : "category",
    title,
    delete: style?.properties["chart:visible"] === "false",
    tickLabelPosition: style?.properties["chart:display-label"] === "false" ? "none" : undefined,
    majorGridlines: childrenNamed(element, "chart:grid").some(
      (grid) => attributeString(grid, "chart:class") !== "minor",
    ),
    minorGridlines: childrenNamed(element, "chart:grid").some(
      (grid) => attributeString(grid, "chart:class") === "minor",
    ),
    shapeProperties: styleShape(style, resources, basePath),
  };
}

function parseTitle(
  element: Element | undefined,
  styleMap: Map<string, ChartStyle>,
  resources: ChartStyleResources,
  basePath: string,
): string | ChartTitleOptions | undefined {
  if (!element) return undefined;
  const text = textOf(childNamed(element, "text:p"));
  if (!text) return undefined;
  const shapeProperties = styleShape(
    requiredStyle(styleMap, attributeString(element, "chart:style-name"), element, basePath),
    resources,
    basePath,
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

function parseWall(
  element: Element | undefined,
  styleMap: Map<string, ChartStyle>,
  resources: ChartStyleResources,
  basePath: string,
) {
  if (!element) return undefined;
  return {
    shapeProperties: styleShape(
      requiredStyle(styleMap, attributeString(element, "chart:style-name"), element, basePath),
      resources,
      basePath,
    ),
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

function chartShape(
  element: Element | undefined,
  styleMap: Map<string, ChartStyle>,
  resources: ChartStyleResources,
  basePath: string,
): ChartShapePropertiesOptions | undefined {
  if (!element) return undefined;
  const styleName = attributeString(element, "chart:style-name");
  if (!styleName) return undefined;
  const style = styleMap.get(styleName);
  if (!style) {
    throw unsupported(
      `${basePath}content.xml`,
      `${chartPath(element)}/@chart:style-name`,
      styleName,
      "unresolved-style-reference",
    );
  }
  return styleShape(style, resources, basePath);
}

function styleShape(
  style: ChartStyle | undefined,
  resources: ChartStyleResources,
  basePath: string,
): ChartShapePropertiesOptions | undefined {
  return style ? projectChartStyle(style, resources, basePath) : undefined;
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
    (node): node is OdfXmlNode => typeof node === "object" && node.name === "table:table-row",
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
                (
                  (cell.children ?? []).find(
                    (child): child is OdfXmlNode =>
                      typeof child === "object" && child.name === "text:p",
                  )?.children ?? []
                ).find((child) => typeof child === "string") ??
                "",
            )
          : "",
      );
    }
  }
  return values;
}

function chartValueText(value: ChartValue | undefined): string {
  if (value === null || value === undefined) return "";
  return typeof value === "object" ? String(value.value) : String(value);
}

function seriesValues(series: ChartSpaceOptions["series"][number]): readonly ChartValue[] {
  return "values" in series ? series.values : series.yValues;
}

function rowCount(table: OdfXmlNode): number {
  return (table.children ?? []).filter(
    (child) => typeof child === "object" && child.name === "table:table-row",
  ).length;
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

function xmlRoot(xml: string): Element {
  const document = parse(xml, { ignoreDeclaration: true, ignoreDoctype: true });
  return document.elements?.[0] ?? document;
}

function emptyChartStyleResources(): ChartStyleResources {
  return {
    styles: [],
    gradients: [],
    hatches: [],
    fillImages: [],
    strokeDashes: [],
    dr3dLights: [],
  };
}

function hasChartStyleResources(resources: ChartStyleResources): boolean {
  return (
    resources.styles.length > 0 ||
    resources.gradients.length > 0 ||
    resources.hatches.length > 0 ||
    resources.fillImages.length > 0 ||
    resources.strokeDashes.length > 0 ||
    (resources.dr3dLights?.length ?? 0) > 0
  );
}
