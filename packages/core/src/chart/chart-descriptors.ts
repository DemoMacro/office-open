/**
 * Chart descriptors — CustomDescriptor for c:chartSpace serialization.
 *
 * @module
 */

import { escapeXml } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { attr, findChild, children } from "@office-open/xml";

import type { CustomDescriptor, ReadContext, WriteContext } from "../descriptor";
import type { ChartSpaceOptions, BubbleSeriesData, ChartSeriesData, ChartType } from "./types";

// ── Mutable build type for read results (readonly properties not assignable) ──

type MutableChartSpaceResult = {
  -readonly [K in keyof ChartSpaceOptions]?: ChartSpaceOptions[K];
};

// ── Chart type → XML element tag mapping ──

const CHART_TYPE_TAGS: Record<string, string> = {
  column: "c:barChart",
  bar: "c:barChart",
  line: "c:lineChart",
  pie: "c:pieChart",
  area: "c:areaChart",
  scatter: "c:scatterChart",
  bubble: "c:bubbleChart",
  doughnut: "c:doughnutChart",
  radar: "c:radarChart",
  stock: "c:stockChart",
  surface: "c:surfaceChart",
};

const CHART_TYPE_TAGS_3D: Record<string, string> = {
  column: "c:bar3DChart",
  bar: "c:bar3DChart",
  line: "c:line3DChart",
  pie: "c:pie3DChart",
  area: "c:area3DChart",
};

const NO_AXES_TYPES = new Set(["pie", "doughnut"]);

// ── XML tag → chart type reverse mapping (for read) ──

const TAG_TO_CHART_TYPE: Record<string, { type: ChartType; threeD?: boolean }> = {
  "c:barChart": { type: "column" },
  "c:bar3DChart": { type: "column", threeD: true },
  "c:lineChart": { type: "line" },
  "c:line3DChart": { type: "line", threeD: true },
  "c:pieChart": { type: "pie" },
  "c:pie3DChart": { type: "pie", threeD: true },
  "c:areaChart": { type: "area" },
  "c:area3DChart": { type: "area", threeD: true },
  "c:scatterChart": { type: "scatter" },
  "c:bubbleChart": { type: "bubble" },
  "c:doughnutChart": { type: "doughnut" },
  "c:radarChart": { type: "radar" },
  "c:stockChart": { type: "stock" },
  "c:surfaceChart": { type: "surface" },
  "c:surface3DChart": { type: "surface", threeD: true },
};

// ── String helpers ──

function attrVal(name: string, value: string | number): string {
  return `${name}="${escapeXml(String(value))}"`;
}

function emptyEl(tag: string): string {
  return `<${tag}/>`;
}

function valEl(tag: string, value: string | number): string {
  return `<${tag} ${attrVal("val", value)}/>`;
}

// ── Series data XML builders ──

function stringifyStrRef(values: readonly string[]): string {
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${escapeXml(v)}</c:v></c:pt>`).join("");
  return `<c:strRef><c:f/><c:strCache><c:ptCount ${attrVal("val", values.length)}/>${pts}</c:strCache></c:strRef>`;
}

function stringifyNumRef(values: readonly number[]): string {
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("");
  return `<c:numRef><c:f/><c:numCache><c:formatCode>General</c:formatCode><c:ptCount ${attrVal("val", values.length)}/>${pts}</c:numCache></c:numRef>`;
}

// ── Chart type header XML ──

function chartTypeHeader(opts: ChartSpaceOptions): string {
  const tag = opts.threeD ? CHART_TYPE_TAGS_3D[opts.type] : CHART_TYPE_TAGS[opts.type];
  if (!tag) throw new Error(`Unsupported chart type: ${opts.type}`);

  const headerParts: string[] = [];

  switch (opts.type) {
    case "column":
    case "bar":
      headerParts.push(valEl("c:barDir", opts.type === "column" ? "col" : "bar"));
      headerParts.push(valEl("c:grouping", "clustered"));
      break;
    case "line":
    case "area":
      headerParts.push(valEl("c:grouping", "standard"));
      break;
    case "scatter":
      headerParts.push(valEl("c:scatterStyle", "line"));
      break;
    case "radar":
      headerParts.push(valEl("c:radarStyle", "standard"));
      break;
    case "pie":
    case "doughnut":
    case "bubble":
      headerParts.push(valEl("c:varyColors", 1));
      break;
  }

  return `<${tag}>${headerParts.join("")}`;
}

function chartTypeFooter(opts: ChartSpaceOptions): string {
  const tag = opts.threeD ? CHART_TYPE_TAGS_3D[opts.type] : CHART_TYPE_TAGS[opts.type];
  // Pie and doughnut have no axId
  if (NO_AXES_TYPES.has(opts.type)) return `</${tag}>`;
  const axIds: string[] = [valEl("c:axId", 10), valEl("c:axId", 20)];
  // 3D charts and surface charts need a third axis
  if (opts.threeD || opts.type === "surface") axIds.push(valEl("c:axId", 30));
  return `${axIds.join("")}</${tag}>`;
}

// ── Series XML ──

function stringifySeries(
  index: number,
  series: ChartSeriesData | BubbleSeriesData,
  categories: readonly string[],
  chartType: ChartType,
): string {
  const parts: string[] = [];

  parts.push(valEl("c:idx", index));
  parts.push(valEl("c:order", index));

  // Series name (c:tx -> c:strRef)
  parts.push(`<c:tx>${stringifyStrRef([series.name])}</c:tx>`);
  parts.push(emptyEl("c:spPr"));

  if (chartType === "bubble") {
    const bs = series as BubbleSeriesData;
    parts.push(`<c:xVal>${stringifyNumRef(bs.xValues)}</c:xVal>`);
    parts.push(`<c:yVal>${stringifyNumRef(bs.yValues)}</c:yVal>`);
    parts.push(`<c:bubbleSize>${stringifyNumRef(bs.bubbleSize)}</c:bubbleSize>`);
  } else if (chartType === "scatter") {
    const s = series as ChartSeriesData;
    parts.push(`<c:xVal>${stringifyStrRef(categories)}</c:xVal>`);
    parts.push(`<c:yVal>${stringifyNumRef(s.values)}</c:yVal>`);
  } else {
    const s = series as ChartSeriesData;
    parts.push(`<c:cat>${stringifyStrRef(categories)}</c:cat>`);
    parts.push(`<c:val>${stringifyNumRef(s.values)}</c:val>`);
  }

  return `<c:ser>${parts.join("")}</c:ser>`;
}

// ── Title XML ──

function stringifyTitle(title: string): string {
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeXml(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
}

// ── Legend XML ──

function stringifyLegend(): string {
  return `<c:legend><c:legendPos val="b"/><c:layout/><c:overlay val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr><c:txPr><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr></c:legend>`;
}

// ── Shared boilerplate ──

function noFillSpPr(): string {
  return `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>`;
}

function chartTxPr(): string {
  return `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
}

// ── Axes XML ──

function stringifyAxes(chartType: ChartType, threeD?: boolean): string {
  if (NO_AXES_TYPES.has(chartType)) return "";

  const parts: string[] = [];

  if (chartType === "scatter" || chartType === "bubble") {
    parts.push(valAxXml(10, 20));
    parts.push(valAxXml(20, 10));
  } else if (chartType === "stock" || chartType === "surface") {
    parts.push(catAxXml(10, 20));
    parts.push(valAxXml(20, 10));
    if (chartType === "surface") {
      parts.push(serAxXml(30, 10));
    }
  } else {
    parts.push(catAxXml(10, 20));
    parts.push(valAxXml(20, 10));
    if (threeD) {
      parts.push(catAxXml(30, 10));
    }
  }

  return parts.join("");
}

function catAxXml(axId: number, crossAx: number): string {
  return `<c:catAx><c:axId ${attrVal("val", axId)}/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx ${attrVal("val", crossAx)}/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
}

function valAxXml(axId: number, crossAx: number): string {
  return `<c:valAx><c:axId ${attrVal("val", axId)}/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:numFmt formatCode="General" sourceLinked="1"/><c:spPr/><c:crossAx ${attrVal("val", crossAx)}/><c:crosses val="autoZero"/></c:valAx>`;
}

function serAxXml(axId: number, crossAx: number): string {
  return `<c:serAx><c:axId ${attrVal("val", axId)}/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:spPr/><c:crossAx ${attrVal("val", crossAx)}/><c:crosses val="autoZero"/><c:tickLblSkip val="1"/></c:serAx>`;
}

// ── Read helpers ──

function readStrCache(el: XmlElement): string[] {
  const strRef = findChild(el, "c:strRef");
  if (!strRef) return [];
  const strCache = findChild(strRef, "c:strCache");
  if (!strCache?.elements) return [];
  const result: string[] = [];
  for (const pt of strCache.elements) {
    if (pt.name === "c:pt" && pt.elements) {
      const v = pt.elements.find((c) => c.name === "c:v");
      if (v?.text !== undefined) result.push(String(v.text));
    }
  }
  return result;
}

function readNumCache(el: XmlElement): number[] {
  const numRef = findChild(el, "c:numRef");
  if (!numRef) return [];
  const numCache = findChild(numRef, "c:numCache");
  if (!numCache?.elements) return [];
  const result: number[] = [];
  for (const pt of numCache.elements) {
    if (pt.name === "c:pt" && pt.elements) {
      const v = pt.elements.find((c) => c.name === "c:v");
      if (v?.text !== undefined) result.push(Number(v.text));
    }
  }
  return result;
}

function readSeriesName(serEl: XmlElement): string {
  const tx = findChild(serEl, "c:tx");
  if (!tx) return "";
  const values = readStrCache(tx);
  return values[0] ?? "";
}

function readTitleText(titleEl: XmlElement): string | undefined {
  const tx = findChild(titleEl, "c:tx");
  if (!tx?.elements) return undefined;
  for (const child of tx.elements) {
    if (child.name === "c:rich" && child.elements) {
      for (const sub of child.elements) {
        if (sub.name === "a:p" && sub.elements) {
          for (const r of sub.elements) {
            if (r.name === "a:r" && r.elements) {
              for (const t of r.elements) {
                if (t.name === "a:t" && t.text !== undefined) {
                  return String(t.text);
                }
              }
            }
          }
        }
      }
    }
  }
  return undefined;
}

// ── Main descriptor ──

export const chartSpaceDesc: CustomDescriptor<ChartSpaceOptions> = {
  kind: "custom",

  stringify(opts: ChartSpaceOptions, _ctx: WriteContext): string {
    const parts: string[] = [];

    // Opening tag with namespaces
    parts.push(
      `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`,
    );

    // Fixed header elements
    parts.push(valEl("c:date1904", 0));
    parts.push(valEl("c:lang", "en-US"));
    parts.push(valEl("c:roundedCorners", 0));

    // Style
    if (opts.style !== undefined) {
      parts.push(valEl("c:style", opts.style));
    }

    // c:chart container
    parts.push("<c:chart>");

    // Title
    if (opts.title) {
      parts.push(stringifyTitle(opts.title));
    }

    parts.push(valEl("c:autoTitleDeleted", 0));

    // c:plotArea
    parts.push("<c:plotArea><c:layout/>");

    // Chart type element (header + series + footer)
    parts.push(chartTypeHeader(opts));

    const categories = opts.categories ?? [];
    for (let i = 0; i < opts.series.length; i++) {
      parts.push(stringifySeries(i, opts.series[i], categories, opts.type));
    }

    parts.push(chartTypeFooter(opts));

    // Axes
    parts.push(stringifyAxes(opts.type, opts.threeD));

    parts.push("</c:plotArea>");

    // Legend
    if (opts.showLegend !== false) {
      parts.push(stringifyLegend());
    }

    parts.push("</c:chart>");

    // SpPr and txPr
    parts.push(noFillSpPr());
    parts.push(chartTxPr());

    parts.push("</c:chartSpace>");

    return parts.join("");
  },

  parse(el: XmlElement, _ctx: ReadContext) {
    const result: MutableChartSpaceResult = {};

    // Style
    const styleEl = findChild(el, "c:style");
    if (styleEl?.attributes?.["val"] !== undefined) {
      result.style = Number(styleEl.attributes["val"]);
    }

    // Chart container
    const chart = findChild(el, "c:chart");
    if (!chart) return result as ChartSpaceOptions;

    // Title
    const titleEl = findChild(chart, "c:title");
    if (titleEl) {
      const title = readTitleText(titleEl);
      if (title !== undefined) result.title = title;
    }

    // Plot area — detect chart type
    const plotArea = findChild(chart, "c:plotArea");
    if (plotArea?.elements) {
      let detectedType: ChartType | undefined;
      let threeD = false;

      for (const child of plotArea.elements) {
        if (!child.name) continue;
        const mapping = TAG_TO_CHART_TYPE[child.name];
        if (mapping) {
          detectedType = mapping.type;
          threeD = mapping.threeD ?? false;
          // bar/column share the c:barChart tag; distinguish via c:barDir
          if (child.name === "c:barChart" || child.name === "c:bar3DChart") {
            const barDir = findChild(child, "c:barDir");
            const barDirVal = barDir ? attr(barDir, "val") : undefined;
            if (barDirVal === "bar") detectedType = "bar";
            else if (barDirVal === "col") detectedType = "column";
          }
          break;
        }
      }

      if (detectedType) {
        result.type = detectedType;
        if (threeD) result.threeD = true;
      }

      // Extract series data
      const seriesEls = children(plotArea, "c:ser");
      if (seriesEls.length > 0) {
        if (detectedType === "bubble") {
          const bubbleSeries: BubbleSeriesData[] = [];
          for (const serEl of seriesEls) {
            const name = readSeriesName(serEl);
            const xVal = findChild(serEl, "c:xVal");
            const yVal = findChild(serEl, "c:yVal");
            const bubbleSize = findChild(serEl, "c:bubbleSize");
            bubbleSeries.push({
              name,
              xValues: xVal ? readNumCache(xVal) : [],
              yValues: yVal ? readNumCache(yVal) : [],
              bubbleSize: bubbleSize ? readNumCache(bubbleSize) : [],
            });
          }
          result.series = bubbleSeries as BubbleSeriesData[];
        } else {
          const chartSeries: ChartSeriesData[] = [];
          let categories: string[] | undefined;

          for (const serEl of seriesEls) {
            const name = readSeriesName(serEl);

            // Read categories from first series
            if (!categories) {
              const catEl = findChild(serEl, "c:cat") ?? findChild(serEl, "c:xVal");
              if (catEl) categories = readStrCache(catEl);
            }

            // Read values
            const valEl = findChild(serEl, "c:val") ?? findChild(serEl, "c:yVal");
            const values = valEl ? readNumCache(valEl) : [];

            chartSeries.push({ name, values });
          }

          result.series = chartSeries as ChartSeriesData[];
          if (categories?.length) result.categories = categories;
        }
      }
    }

    // Legend — default is true, only set if explicitly absent
    const legend = findChild(chart, "c:legend");
    if (!legend) {
      result.showLegend = false;
    }

    return result as ChartSpaceOptions;
  },
};
