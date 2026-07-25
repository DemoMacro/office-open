/**
 * Chart descriptors — CustomDescriptor for c:chartSpace serialization.
 *
 * @module
 */

import { escapeXml } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { attr, findChild, children, textOf } from "@office-open/xml";

import type { CustomDescriptor, ReadContext, WriteContext } from "../descriptor";
import type {
  ChartSpaceOptions,
  BubbleSeriesData,
  ChartSeriesData,
  ChartType,
  DataLabelOptions,
  DataLabelsOptions,
  ErrorBarOptions,
  TrendlineLabelOptions,
  TrendlineOptions,
  View3DOptions,
} from "./types";

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

// ── Boolean helper (CT_OnOff: true → omit val, false → "0") ──

function boolVal(value: boolean | undefined): string {
  // XSD CT_OnOff: omit val for true, val="0" for false
  return value === false ? ` val="0"` : "";
}

// ── Numeric literal (c:numLit) ──

function stringifyNumLit(value: number): string {
  return `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="1"/><c:pt idx="0"><c:v>${value}</c:v></c:pt></c:numLit>`;
}

// ── Trendline XML (CT_Trendline) ──

function stringifyTrendline(opts: TrendlineOptions): string {
  const parts: string[] = [];
  if (opts.name !== undefined) {
    parts.push(`<c:name>${escapeXml(opts.name)}</c:name>`);
  }
  parts.push(valEl("c:trendlineType", opts.type ?? "linear"));
  if (opts.order !== undefined) parts.push(valEl("c:order", opts.order));
  if (opts.period !== undefined) parts.push(valEl("c:period", opts.period));
  if (opts.forward !== undefined) parts.push(valEl("c:forward", opts.forward));
  if (opts.backward !== undefined) parts.push(valEl("c:backward", opts.backward));
  if (opts.intercept !== undefined) parts.push(valEl("c:intercept", opts.intercept));
  if (opts.dispRSqr !== undefined) parts.push(`<c:dispRSqr${boolVal(opts.dispRSqr)}/>`);
  if (opts.dispEq !== undefined) parts.push(`<c:dispEq${boolVal(opts.dispEq)}/>`);
  // XSD CT_Trendline: trendlineLbl follows dispEq; CT_TrendlineLbl's children are all optional.
  const labelFmt = opts.label?.numberFormat;
  if (labelFmt !== undefined) {
    parts.push(
      `<c:trendlineLbl><c:numFmt formatCode="${escapeXml(labelFmt)}" sourceLinked="0"/></c:trendlineLbl>`,
    );
  }
  return `<c:trendline>${parts.join("")}</c:trendline>`;
}

// ── Error bars XML (CT_ErrBars) ──

function stringifyErrBars(opts: ErrorBarOptions): string {
  const parts: string[] = [];
  if (opts.direction !== undefined) parts.push(valEl("c:errDir", opts.direction));
  parts.push(valEl("c:errBarType", opts.barType ?? "both"));
  parts.push(valEl("c:errValType", opts.valueType ?? "fixedVal"));
  if (opts.noEndCap !== undefined) parts.push(`<c:noEndCap${boolVal(opts.noEndCap)}/>`);
  if (opts.plusValue !== undefined)
    parts.push(`<c:plus>${stringifyNumLit(opts.plusValue)}</c:plus>`);
  if (opts.minusValue !== undefined)
    parts.push(`<c:minus>${stringifyNumLit(opts.minusValue)}</c:minus>`);
  if (opts.value !== undefined) parts.push(valEl("c:val", opts.value));
  return `<c:errBars>${parts.join("")}</c:errBars>`;
}

// ── Data labels XML (CT_DLbls) ──

function stringifyDataLabel(opts: DataLabelOptions): string {
  // CT_DLbl: idx (required) → choice(delete | Group_DLbl needing ≥1 EG_DLblShared child).
  if (opts.delete) {
    return `<c:dLbl><c:idx val="${opts.index}"/><c:delete val="1"/></c:dLbl>`;
  }
  const inner: string[] = [];
  if (opts.numberFormat !== undefined)
    inner.push(`<c:numFmt formatCode="${escapeXml(opts.numberFormat)}" sourceLinked="0"/>`);
  if (opts.position !== undefined) inner.push(valEl("c:dLblPos", opts.position));
  return `<c:dLbl><c:idx val="${opts.index}"/>${inner.join("")}</c:dLbl>`;
}

function stringifyDataLabels(opts: DataLabelsOptions): string {
  const parts: string[] = [];
  // CT_DLbls: per-point overrides (c:dLbl) precede the shared Group_DLbls settings.
  if (opts.labels) {
    for (const lbl of opts.labels) parts.push(stringifyDataLabel(lbl));
  }
  if (opts.position !== undefined) parts.push(valEl("c:dLblPos", opts.position));
  if (opts.showLegendKey !== undefined)
    parts.push(`<c:showLegendKey${boolVal(opts.showLegendKey)}/>`);
  if (opts.showVal !== undefined) parts.push(`<c:showVal${boolVal(opts.showVal)}/>`);
  if (opts.showCatName !== undefined) parts.push(`<c:showCatName${boolVal(opts.showCatName)}/>`);
  if (opts.showSerName !== undefined) parts.push(`<c:showSerName${boolVal(opts.showSerName)}/>`);
  if (opts.showPercent !== undefined) parts.push(`<c:showPercent${boolVal(opts.showPercent)}/>`);
  if (opts.showBubbleSize !== undefined)
    parts.push(`<c:showBubbleSize${boolVal(opts.showBubbleSize)}/>`);
  if (opts.separator !== undefined)
    parts.push(`<c:separator>${escapeXml(opts.separator)}</c:separator>`);
  if (opts.showLeaderLines !== undefined)
    parts.push(`<c:showLeaderLines${boolVal(opts.showLeaderLines)}/>`);
  // CT_ChartLines (leaderLines) closes Group_DLbls; an empty element is XSD-valid.
  if (opts.leaderLines) parts.push(emptyEl("c:leaderLines"));
  return `<c:dLbls>${parts.join("")}</c:dLbls>`;
}

// ── 3D view XML (CT_View3D) ──

function stringifyView3D(opts: View3DOptions): string {
  const parts: string[] = [];
  if (opts.rotX !== undefined) parts.push(valEl("c:rotX", opts.rotX));
  if (opts.hPercent !== undefined) parts.push(valEl("c:hPercent", opts.hPercent));
  if (opts.rotY !== undefined) parts.push(valEl("c:rotY", opts.rotY));
  if (opts.depthPercent !== undefined) parts.push(valEl("c:depthPercent", opts.depthPercent));
  if (opts.rAngAx !== undefined) parts.push(`<c:rAngAx${boolVal(opts.rAngAx)}/>`);
  if (opts.perspective !== undefined) parts.push(valEl("c:perspective", opts.perspective));
  return `<c:view3D>${parts.join("")}</c:view3D>`;
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

  // Per-series enhancements (XSD order: dLbls → trendline → errBars → cat/val)
  const s = series as ChartSeriesData;
  if (s.dataLabels) parts.push(stringifyDataLabels(s.dataLabels));
  if (s.trendlines) {
    for (const tl of s.trendlines) parts.push(stringifyTrendline(tl));
  }
  if (s.errorBars) parts.push(stringifyErrBars(s.errorBars));

  if (chartType === "bubble") {
    const bs = series as BubbleSeriesData;
    parts.push(`<c:xVal>${stringifyNumRef(bs.xValues)}</c:xVal>`);
    parts.push(`<c:yVal>${stringifyNumRef(bs.yValues)}</c:yVal>`);
    parts.push(`<c:bubbleSize>${stringifyNumRef(bs.bubbleSize)}</c:bubbleSize>`);
  } else if (chartType === "scatter") {
    parts.push(`<c:xVal>${stringifyStrRef(categories)}</c:xVal>`);
    parts.push(`<c:yVal>${stringifyNumRef(s.values)}</c:yVal>`);
  } else {
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

// ── Read helpers: series enhancements ──

function readBoolAttr(el: XmlElement, childName: string): boolean | undefined {
  const child = findChild(el, childName);
  if (!child) return undefined;
  const v = attr(child, "val");
  // CT_OnOff: absent val or val="1"/"true" → true; val="0"/"false" → false
  if (v === undefined) return true;
  return v !== "0" && v !== "false";
}

function readValNum(el: XmlElement, childName: string): number | undefined {
  const child = findChild(el, childName);
  if (!child) return undefined;
  const v = attr(child, "val");
  return v !== undefined ? Number(v) : undefined;
}

function readValStr(el: XmlElement, childName: string): string | undefined {
  const child = findChild(el, childName);
  if (!child) return undefined;
  return attr(child, "val") ?? undefined;
}

function readNumLit(el: XmlElement, childName: string): number | undefined {
  const child = findChild(el, childName);
  if (!child) return undefined;
  const numLit = findChild(child, "c:numLit");
  if (!numLit?.elements) return undefined;
  for (const pt of numLit.elements) {
    if (pt.name === "c:pt" && pt.elements) {
      const v = pt.elements.find((c) => c.name === "c:v");
      if (v) {
        const text = textOf(v);
        if (text) return Number(text);
      }
    }
  }
  return undefined;
}

function readTrendlines(serEl: XmlElement): TrendlineOptions[] | undefined {
  const tlEls = children(serEl, "c:trendline");
  if (tlEls.length === 0) return undefined;
  return tlEls.map((tl) => {
    const opts: TrendlineOptions = {};
    const type = readValStr(tl, "c:trendlineType");
    if (type) opts.type = type as TrendlineOptions["type"];
    const order = readValNum(tl, "c:order");
    if (order !== undefined) opts.order = order;
    const period = readValNum(tl, "c:period");
    if (period !== undefined) opts.period = period;
    const forward = readValNum(tl, "c:forward");
    if (forward !== undefined) opts.forward = forward;
    const backward = readValNum(tl, "c:backward");
    if (backward !== undefined) opts.backward = backward;
    const intercept = readValNum(tl, "c:intercept");
    if (intercept !== undefined) opts.intercept = intercept;
    const dispRSqr = readBoolAttr(tl, "c:dispRSqr");
    if (dispRSqr !== undefined) opts.dispRSqr = dispRSqr;
    const dispEq = readBoolAttr(tl, "c:dispEq");
    if (dispEq !== undefined) opts.dispEq = dispEq;
    const nameEl = findChild(tl, "c:name");
    const nameText = textOf(nameEl);
    if (nameText) opts.name = nameText;
    const labelFmt = attr(findChild(findChild(tl, "c:trendlineLbl"), "c:numFmt"), "formatCode");
    if (labelFmt) opts.label = { numberFormat: labelFmt } satisfies TrendlineLabelOptions;
    return opts;
  });
}

function readErrBars(serEl: XmlElement): ErrorBarOptions | undefined {
  const ebEl = findChild(serEl, "c:errBars");
  if (!ebEl) return undefined;
  const opts: ErrorBarOptions = {};
  const direction = readValStr(ebEl, "c:errDir");
  if (direction) opts.direction = direction as ErrorBarOptions["direction"];
  const barType = readValStr(ebEl, "c:errBarType");
  if (barType) opts.barType = barType as ErrorBarOptions["barType"];
  const valueType = readValStr(ebEl, "c:errValType");
  if (valueType) opts.valueType = valueType as ErrorBarOptions["valueType"];
  const noEndCap = readBoolAttr(ebEl, "c:noEndCap");
  if (noEndCap !== undefined) opts.noEndCap = noEndCap;
  const plusValue = readNumLit(ebEl, "c:plus");
  if (plusValue !== undefined) opts.plusValue = plusValue;
  const minusValue = readNumLit(ebEl, "c:minus");
  if (minusValue !== undefined) opts.minusValue = minusValue;
  const value = readValNum(ebEl, "c:val");
  if (value !== undefined) opts.value = value;
  return opts;
}

function readDataLabels(serEl: XmlElement): DataLabelsOptions | undefined {
  const dlEl = findChild(serEl, "c:dLbls");
  if (!dlEl) return undefined;
  const opts: DataLabelsOptions = {};
  const position = readValStr(dlEl, "c:dLblPos");
  if (position) opts.position = position as DataLabelsOptions["position"];
  const showLegendKey = readBoolAttr(dlEl, "c:showLegendKey");
  if (showLegendKey !== undefined) opts.showLegendKey = showLegendKey;
  const showVal = readBoolAttr(dlEl, "c:showVal");
  if (showVal !== undefined) opts.showVal = showVal;
  const showCatName = readBoolAttr(dlEl, "c:showCatName");
  if (showCatName !== undefined) opts.showCatName = showCatName;
  const showSerName = readBoolAttr(dlEl, "c:showSerName");
  if (showSerName !== undefined) opts.showSerName = showSerName;
  const showPercent = readBoolAttr(dlEl, "c:showPercent");
  if (showPercent !== undefined) opts.showPercent = showPercent;
  const showBubbleSize = readBoolAttr(dlEl, "c:showBubbleSize");
  if (showBubbleSize !== undefined) opts.showBubbleSize = showBubbleSize;
  const showLeaderLines = readBoolAttr(dlEl, "c:showLeaderLines");
  if (showLeaderLines !== undefined) opts.showLeaderLines = showLeaderLines;
  const sepEl = findChild(dlEl, "c:separator");
  const sepText = textOf(sepEl);
  if (sepText) opts.separator = sepText;
  const lblEls = children(dlEl, "c:dLbl");
  if (lblEls.length > 0) {
    opts.labels = lblEls.map((el): DataLabelOptions => {
      const result: DataLabelOptions = { index: readValNum(el, "c:idx") ?? 0 };
      if (findChild(el, "c:delete")) result.delete = true;
      const numFmt = attr(findChild(el, "c:numFmt"), "formatCode");
      if (numFmt) result.numberFormat = numFmt;
      const pos = readValStr(el, "c:dLblPos");
      if (pos) result.position = pos as DataLabelOptions["position"];
      return result;
    });
  }
  if (findChild(dlEl, "c:leaderLines")) opts.leaderLines = true;
  return opts;
}

function readView3D(chartEl: XmlElement): View3DOptions | undefined {
  const v3d = findChild(chartEl, "c:view3D");
  if (!v3d) return undefined;
  const opts: View3DOptions = {};
  const rotX = readValNum(v3d, "c:rotX");
  if (rotX !== undefined) opts.rotX = rotX;
  const hPercent = readValNum(v3d, "c:hPercent");
  if (hPercent !== undefined) opts.hPercent = hPercent;
  const rotY = readValNum(v3d, "c:rotY");
  if (rotY !== undefined) opts.rotY = rotY;
  const depthPercent = readValNum(v3d, "c:depthPercent");
  if (depthPercent !== undefined) opts.depthPercent = depthPercent;
  const rAngAx = readBoolAttr(v3d, "c:rAngAx");
  if (rAngAx !== undefined) opts.rAngAx = rAngAx;
  const perspective = readValNum(v3d, "c:perspective");
  if (perspective !== undefined) opts.perspective = perspective;
  return opts;
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

    // 3D view (before plotArea per CT_Chart sequence)
    if (opts.view3D) {
      parts.push(stringifyView3D(opts.view3D));
    }

    // c:plotArea
    parts.push("<c:plotArea><c:layout/>");

    // Chart type element (header + series + footer)
    parts.push(chartTypeHeader(opts));

    const categories = opts.categories ?? [];
    for (const [i, series] of opts.series.entries()) {
      parts.push(stringifySeries(i, series, categories, opts.type));
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

    // 3D view (before plotArea per CT_Chart sequence)
    const view3D = readView3D(chart);
    if (view3D) result.view3D = view3D;

    // Plot area — detect chart type
    const plotArea = findChild(chart, "c:plotArea");
    if (plotArea?.elements) {
      let detectedType: ChartType | undefined;
      let threeD = false;
      let chartTypeEl: XmlElement | undefined;

      for (const child of plotArea.elements) {
        if (!child.name) continue;
        const mapping = TAG_TO_CHART_TYPE[child.name];
        if (mapping) {
          detectedType = mapping.type;
          threeD = mapping.threeD ?? false;
          chartTypeEl = child;
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

      // Extract series data (c:ser are children of the chart-type element, not plotArea)
      const seriesEls = chartTypeEl ? children(chartTypeEl, "c:ser") : [];
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

            const ser: ChartSeriesData = { name, values };

            // Read series enhancements
            const trendlines = readTrendlines(serEl);
            if (trendlines) ser.trendlines = trendlines;
            const errorBars = readErrBars(serEl);
            if (errorBars) ser.errorBars = errorBars;
            const dataLabels = readDataLabels(serEl);
            if (dataLabels) ser.dataLabels = dataLabels;

            chartSeries.push(ser);
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
