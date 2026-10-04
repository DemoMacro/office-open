import type { Element } from "@office-open/xml";

import { ODF_NAMESPACES, escapeText, metaXml, parseMeta } from "./meta";
import { parseOdfNodes, serializeOdfNodes, type OdfXmlNode } from "./odf-node";
import { generateOcf, readOcf, readXml } from "./package";
import {
  attributeString,
  childNamed,
  childrenNamed,
  emuToLength,
  lengthToEmu,
  textOf,
  xmlElement,
} from "./xml";

const MIME = "application/vnd.oasis.opendocument.chart";
const NAMESPACES = `${ODF_NAMESPACES} xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"`;
const childAttribute = attributeString;

/** Root chart:chart element of an embedded chart subdocument. */
export interface ChartChartOptions {
  /** Chart type as a namespaced token, for example chart:bar. */
  class: string;
  width?: number;
  height?: number;
  href?: string;
  styleName?: string;
  title?: ChartTitleOptions;
  subtitle?: ChartTitleOptions;
  footer?: ChartTitleOptions;
  legend?: ChartLegendOptions;
  plotArea: ChartPlotAreaOptions;
  /** table:table child preserved as generic ODF nodes for cell-source round trips. */
  table?: OdfXmlNode[];
}

export interface ChartDocumentOptions {
  title?: string;
  chart: ChartChartOptions;
}

/** chart:title, chart:subtitle, or chart:footer. */
export interface ChartTitleOptions {
  text?: string;
  cellRange?: string;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  styleName?: string;
}

interface ChartLegendBase {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  expansion?: "wide" | "high" | "balanced" | "custom";
  expansionAspectRatio?: number;
  styleName?: string;
}

export interface ChartStandardLegendOptions extends ChartLegendBase {
  position: "start" | "end" | "top" | "bottom";
  align?: "start" | "center" | "end";
}

export interface ChartCornerLegendOptions extends ChartLegendBase {
  position: "top-start" | "bottom-start" | "top-end" | "bottom-end";
}

export type ChartLegendOptions = ChartStandardLegendOptions | ChartCornerLegendOptions;

export interface ChartPlotAreaOptions {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  styleName?: string;
  cellRange?: string;
  dataSourceHasLabels?: "none" | "row" | "column" | "both";
  axes?: ChartAxisOptions[];
  series?: ChartSeriesOptions[];
  wall?: ChartWallOptions;
  floor?: ChartWallOptions;
}

export interface ChartAxisOptions {
  dimension: "x" | "y" | "z";
  name?: string;
  styleName?: string;
  title?: ChartTitleOptions;
  categories?: string;
  grids?: ChartGridOptions[];
}

export interface ChartGridOptions {
  class?: "major" | "minor";
  styleName?: string;
}

export interface ChartSeriesOptions {
  values?: string;
  label?: string;
  class?: string;
  attachedAxis?: string;
  styleName?: string;
  domains?: string[];
  dataPoints?: ChartDataPointOptions[];
}

export interface ChartDataPointOptions {
  repeated?: number;
  styleName?: string;
}

export interface ChartWallOptions {
  width?: number;
  styleName?: string;
}

export function generateChartDocument(options: ChartDocumentOptions): Uint8Array {
  return generateOcf(MIME, {
    "content.xml": chartContentXml(chartXml(options.chart)),
    "meta.xml": metaXml({ title: options.title }),
  });
}

export function parseChartDocument(data: Uint8Array): ChartDocumentOptions {
  const { files } = readOcf(data, MIME);
  const content = readXml(files, "content.xml");
  const body = childNamed(childNamed(content, "office:body"), "office:chart");
  return { ...parseMeta(files), chart: parseChart(childNamed(body, "chart:chart")) };
}

function chartContentXml(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.3"><office:body><office:chart>${body}</office:chart></office:body></office:document-content>`;
}

function chartXml(chart: ChartChartOptions): string {
  return xmlElement(
    "chart:chart",
    {
      "chart:class": chart.class,
      "svg:width": odfLength(chart.width),
      "svg:height": odfLength(chart.height),
      "xlink:href": chart.href,
      "xlink:type": chart.href ? "simple" : undefined,
      "chart:style-name": chart.styleName,
    },
    [
      titleXml("chart:title", chart.title),
      titleXml("chart:subtitle", chart.subtitle),
      titleXml("chart:footer", chart.footer),
      legendXml(chart.legend),
      plotAreaXml(chart.plotArea),
      ...serializeOdfNodes(chart.table),
    ],
  );
}

function titleXml(name: string, title: ChartTitleOptions | undefined): string {
  if (!title) return "";
  return xmlElement(
    name,
    titleAttributes(title),
    title.text ? [xmlElement("text:p", undefined, [escapeText(title.text)])] : [],
  );
}

function titleAttributes(title: ChartTitleOptions) {
  return {
    "table:cell-range": title.cellRange,
    "svg:x": odfLength(title.x),
    "svg:y": odfLength(title.y),
    "svg:width": odfLength(title.width),
    "svg:height": odfLength(title.height),
    "chart:style-name": title.styleName,
  };
}

function legendXml(legend: ChartLegendOptions | undefined): string {
  if (!legend) return "";
  const corner = legend.position.includes("-");
  const align = "align" in legend ? legend.align : undefined;
  return xmlElement("chart:legend", {
    "chart:legend-position": legend.position,
    "chart:legend-align": corner ? undefined : align,
    "svg:x": odfLength(legend.x),
    "svg:y": odfLength(legend.y),
    "svg:width": odfLength(legend.width),
    "svg:height": odfLength(legend.height),
    "style:legend-expansion": legend.expansion,
    "style:legend-expansion-aspect-ratio":
      legend.expansion === "custom" ? legend.expansionAspectRatio : undefined,
    "chart:style-name": legend.styleName,
  });
}

function plotAreaXml(plotArea: ChartPlotAreaOptions): string {
  return xmlElement(
    "chart:plot-area",
    {
      "svg:x": odfLength(plotArea.x),
      "svg:y": odfLength(plotArea.y),
      "svg:width": odfLength(plotArea.width),
      "svg:height": odfLength(plotArea.height),
      "chart:style-name": plotArea.styleName,
      "table:cell-range-address": plotArea.cellRange,
      "chart:data-source-has-labels": plotArea.dataSourceHasLabels,
    },
    [
      ...(plotArea.axes ?? []).map(axisXml),
      ...(plotArea.series ?? []).map(seriesXml),
      plotArea.wall
        ? xmlElement("chart:wall", {
            "svg:width": odfLength(plotArea.wall.width),
            "chart:style-name": plotArea.wall.styleName,
          })
        : "",
      plotArea.floor
        ? xmlElement("chart:floor", {
            "svg:width": odfLength(plotArea.floor.width),
            "chart:style-name": plotArea.floor.styleName,
          })
        : "",
    ],
  );
}

function axisXml(axis: ChartAxisOptions): string {
  return xmlElement(
    "chart:axis",
    {
      "chart:dimension": axis.dimension,
      "chart:name": axis.name,
      "chart:style-name": axis.styleName,
    },
    [
      titleXml("chart:title", axis.title),
      axis.categories
        ? xmlElement("chart:categories", { "table:cell-range-address": axis.categories })
        : "",
      ...(axis.grids ?? []).map((grid) =>
        xmlElement("chart:grid", {
          "chart:class": grid.class,
          "chart:style-name": grid.styleName,
        }),
      ),
    ],
  );
}

function seriesXml(series: ChartSeriesOptions): string {
  return xmlElement(
    "chart:series",
    {
      "chart:values-cell-range-address": series.values,
      "chart:label-cell-address": series.label,
      "chart:class": series.class,
      "chart:attached-axis": series.attachedAxis,
      "chart:style-name": series.styleName,
    },
    [
      ...(series.domains ?? []).map((domain) =>
        xmlElement("chart:domain", { "table:cell-range-address": domain }),
      ),
      ...(series.dataPoints ?? []).map((point) =>
        xmlElement("chart:data-point", {
          "chart:repeated": point.repeated,
          "chart:style-name": point.styleName,
        }),
      ),
    ],
  );
}

function parseChart(element: Element | undefined): ChartChartOptions {
  const title = (name: string) => parseTitle(childNamed(element, name));
  return {
    class: childAttribute(element, "chart:class")!,
    width: lengthToEmu(childAttribute(element, "svg:width")),
    height: lengthToEmu(childAttribute(element, "svg:height")),
    href: childAttribute(element, "xlink:href"),
    styleName: childAttribute(element, "chart:style-name"),
    title: title("chart:title"),
    subtitle: title("chart:subtitle"),
    footer: title("chart:footer"),
    legend: parseLegend(childNamed(element, "chart:legend")),
    plotArea: parsePlotArea(childNamed(element, "chart:plot-area")),
    table: parseOdfNodes(element).filter((node) => node.name === "table:table"),
  };
}

function parseTitle(element: Element | undefined): ChartTitleOptions | undefined {
  if (!element) return undefined;
  return {
    text: textOf(childNamed(element, "text:p")) || undefined,
    cellRange: childAttribute(element, "table:cell-range"),
    x: lengthToEmu(childAttribute(element, "svg:x")),
    y: lengthToEmu(childAttribute(element, "svg:y")),
    width: lengthToEmu(childAttribute(element, "svg:width")),
    height: lengthToEmu(childAttribute(element, "svg:height")),
    styleName: childAttribute(element, "chart:style-name"),
  };
}

function parseLegend(element: Element | undefined): ChartLegendOptions | undefined {
  if (!element) return undefined;
  return {
    position: childAttribute(element, "chart:legend-position")! as ChartLegendOptions["position"],
    align: childAttribute(element, "chart:legend-align") as ChartStandardLegendOptions["align"],
    x: lengthToEmu(childAttribute(element, "svg:x")),
    y: lengthToEmu(childAttribute(element, "svg:y")),
    width: lengthToEmu(childAttribute(element, "svg:width")),
    height: lengthToEmu(childAttribute(element, "svg:height")),
    expansion: childAttribute(element, "style:legend-expansion") as ChartLegendOptions["expansion"],
    expansionAspectRatio:
      Number(childAttribute(element, "style:legend-expansion-aspect-ratio") ?? Number.NaN) ||
      undefined,
    styleName: childAttribute(element, "chart:style-name"),
  };
}

function parsePlotArea(element: Element | undefined): ChartPlotAreaOptions {
  return {
    x: lengthToEmu(childAttribute(element, "svg:x")),
    y: lengthToEmu(childAttribute(element, "svg:y")),
    width: lengthToEmu(childAttribute(element, "svg:width")),
    height: lengthToEmu(childAttribute(element, "svg:height")),
    styleName: childAttribute(element, "chart:style-name"),
    cellRange: childAttribute(element, "table:cell-range-address"),
    dataSourceHasLabels: childAttribute(
      element,
      "chart:data-source-has-labels",
    ) as ChartPlotAreaOptions["dataSourceHasLabels"],
    axes: childrenNamed(element, "chart:axis").map(parseAxis),
    series: childrenNamed(element, "chart:series").map(parseSeries),
    wall: parseWall(childNamed(element, "chart:wall")),
    floor: parseWall(childNamed(element, "chart:floor")),
  };
}

function parseAxis(element: Element): ChartAxisOptions {
  return {
    dimension: childAttribute(element, "chart:dimension")! as ChartAxisOptions["dimension"],
    name: childAttribute(element, "chart:name"),
    styleName: childAttribute(element, "chart:style-name"),
    title: parseTitle(childNamed(element, "chart:title")),
    categories: childAttribute(childNamed(element, "chart:categories"), "table:cell-range-address"),
    grids: childrenNamed(element, "chart:grid").map((grid) => ({
      class: childAttribute(grid, "chart:class") as ChartGridOptions["class"],
      styleName: childAttribute(grid, "chart:style-name"),
    })),
  };
}

function parseSeries(element: Element): ChartSeriesOptions {
  return {
    values: childAttribute(element, "chart:values-cell-range-address"),
    label: childAttribute(element, "chart:label-cell-address"),
    class: childAttribute(element, "chart:class"),
    attachedAxis: childAttribute(element, "chart:attached-axis"),
    styleName: childAttribute(element, "chart:style-name"),
    domains: childrenNamed(element, "chart:domain")
      .map((domain) => childAttribute(domain, "table:cell-range-address"))
      .filter((domain) => domain !== undefined),
    dataPoints: childrenNamed(element, "chart:data-point").map((point) => ({
      repeated: Number(childAttribute(point, "chart:repeated") ?? Number.NaN) || undefined,
      styleName: childAttribute(point, "chart:style-name"),
    })),
  };
}

function parseWall(element: Element | undefined): ChartWallOptions | undefined {
  if (!element) return undefined;
  return {
    width: lengthToEmu(childAttribute(element, "svg:width")),
    styleName: childAttribute(element, "chart:style-name"),
  };
}

/** EMU number → ODF length; undefined stays absent. */
function odfLength(value: number | undefined): string | undefined {
  return value === undefined ? undefined : emuToLength(value);
}
