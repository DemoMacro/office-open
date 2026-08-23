/**
 * Chart descriptors — CustomDescriptor for c:chartSpace serialization.
 *
 * @module
 */

import { escapeXml, stringify as stringifyXml, stringifyElement } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { attr, findChild, findFirst, children, textOf } from "@office-open/xml";

import type { CustomDescriptor, ReadContext, WriteContext } from "../descriptor";
import { shapePropertiesDesc } from "../drawing/shape-properties-desc";
import type { ShapePropertiesOptions } from "../drawing/shape-properties-desc";
import { textBodyDesc } from "../drawing/text/text-body";
import type { TextBodyOptions } from "../drawing/text/text-body";
import { parseColorMapping, stringifyColorMapping } from "../theme/color-mapping";
import {
  xsdAxisCrossBetween,
  xsdAxisCrosses,
  xsdAxisLabelAlignment,
  xsdAxisOrientation,
  xsdAxisPosition,
  xsdDataLabelPosition,
  xsdErrorValueType,
  xsdLegendPosition,
  xsdSizeRepresents,
  xsdSplitType,
  xsdTrendlineType,
} from "../util/mappings";
import { parseOnOff } from "../util/values";
import type {
  ChartSpaceOptions,
  BubbleSeriesData,
  ChartSeriesData,
  ChartLinesOptions,
  ChartGrouping,
  SecondaryChartGroupOptions,
  ChartType,
  DataLabelOptions,
  DataLabelsOptions,
  ErrorBarOptions,
  TrendlineLabelOptions,
  TrendlineOptions,
  View3DOptions,
  AxisOptions,
  AxisKind,
  AxisScalingOptions,
  AxisPosition,
  AxisTickMark,
  AxisTickLabelPosition,
  AxisCrosses,
  AxisCrossBetween,
  AxisOrientation,
  AxisLabelAlignment,
  BuiltInDisplayUnit,
  DisplayUnitsOptions,
  TimeUnit,
  MarkerOptions,
  MarkerSymbol,
  DataPointOptions,
  PictureOptionsOptions,
  PictureFormat,
  BarShape,
  ChartSeriesCommon,
  ScatterSeriesData,
  ChartTitleOptions,
  ManualLayoutOptions,
  SurfaceOptions,
  LayoutTarget,
  LayoutMode,
  DisplayBlanksAs,
  SizeRepresents,
  DataTableOptions,
  OfPieType,
  SplitType,
  ProtectionOptions,
  ExternalDataOptions,
  ChartHeaderFooterOptions,
  ChartPageMarginsOptions,
  ChartPageSetupOptions,
  PrintSettingsOptions,
  PageSetupOrientation,
  PivotSourceOptions,
  ChartPivotFormatOptions,
  BandFormatOptions,
  LegendEntryOptions,
  RadarStyle,
  LegendPosition,
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
  ofPie: "c:ofPieChart",
};

const CHART_TYPE_TAGS_3D: Record<string, string> = {
  column: "c:bar3DChart",
  bar: "c:bar3DChart",
  line: "c:line3DChart",
  pie: "c:pie3DChart",
  area: "c:area3DChart",
};

const NO_AXES_TYPES = new Set(["pie", "doughnut", "ofPie"]);

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
  "c:ofPieChart": { type: "ofPie" },
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

// ── Boolean helper (CT_OnOff) ──

function boolVal(value: boolean | undefined): string {
  // Office chart writers always emit the explicit val attribute (val="1"/"0");
  // the XSD allows omitting val for true, but the corpus never does.
  return ` val="${value === false ? 0 : 1}"`;
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
  parts.push(valEl("c:trendlineType", xsdTrendlineType.to(opts.type ?? "linear")));
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
  // ST_ErrDir is x|y — skip values outside the enum (a legacy "both" belongs
  // to errBarType, not errDir)
  if (opts.direction === "x" || opts.direction === "y") {
    parts.push(valEl("c:errDir", opts.direction));
  }
  parts.push(valEl("c:errBarType", opts.barType ?? "both"));
  parts.push(valEl("c:errValType", xsdErrorValueType.to(opts.valueType ?? "fixedValue")));
  if (opts.noEndCap !== undefined) parts.push(`<c:noEndCap${boolVal(opts.noEndCap)}/>`);
  if (opts.plusValue !== undefined)
    parts.push(`<c:plus>${stringifyNumLit(opts.plusValue)}</c:plus>`);
  if (opts.minusValue !== undefined)
    parts.push(`<c:minus>${stringifyNumLit(opts.minusValue)}</c:minus>`);
  if (opts.value !== undefined) parts.push(valEl("c:val", opts.value));
  return `<c:errBars>${parts.join("")}</c:errBars>`;
}

// ── Data labels XML (CT_DLbls) ──

/** EG_DLblShared members — emitted identically by per-point c:dLbl and group c:dLbls. */
function dataLabelSharedParts(
  opts: DataLabelOptions | DataLabelsOptions,
  ctx: WriteContext,
): string[] {
  const parts: string[] = [];
  if (opts.numberFormat !== undefined)
    parts.push(`<c:numFmt formatCode="${escapeXml(opts.numberFormat)}" sourceLinked="0"/>`);
  parts.push(chartSpPr(opts.shapeProperties, ctx));
  if (opts.textProperties)
    parts.push(`<c:txPr>${textBodyDesc.stringify(opts.textProperties, ctx) ?? ""}</c:txPr>`);
  if (opts.position !== undefined)
    parts.push(valEl("c:dLblPos", xsdDataLabelPosition.to(opts.position)));
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
  return parts;
}

function stringifyDataLabel(opts: DataLabelOptions, ctx: WriteContext): string {
  // CT_DLbl: idx (required) → choice(delete | Group_DLbl needing ≥1 EG_DLblShared child).
  if (opts.delete) {
    return `<c:dLbl><c:idx val="${opts.index}"/><c:delete val="1"/></c:dLbl>`;
  }
  const inner: string[] = [];
  // Group_DLbl: layout → tx → EG_DLblShared (numFmt…dLblPos…show*…separator).
  if (opts.layout === true) inner.push("<c:layout/>");
  else if (opts.layout) inner.push(stringifyLayout(opts.layout));
  if (opts.text)
    inner.push(`<c:tx><c:rich>${textBodyDesc.stringify(opts.text, ctx) ?? ""}</c:rich></c:tx>`);
  inner.push(...dataLabelSharedParts(opts, ctx));
  return `<c:dLbl><c:idx val="${opts.index}"/>${inner.join("")}</c:dLbl>`;
}

// Group_DLbls members — present in any of these, the choice resolves to the
// shared-settings arm (CT_DLbls is choice(c:delete | Group_DLbls))
const DATA_LABELS_SHARED_KEYS = [
  "numberFormat",
  "shapeProperties",
  "textProperties",
  "position",
  "showLegendKey",
  "showVal",
  "showCatName",
  "showSerName",
  "showPercent",
  "showBubbleSize",
  "separator",
  "showLeaderLines",
  "leaderLines",
] as const;

function stringifyDataLabels(opts: DataLabelsOptions, ctx: WriteContext): string {
  const head = opts.labels?.map((lbl) => stringifyDataLabel(lbl, ctx)).join("") ?? "";
  // CT_DLbls choice arms: c:dLbl overrides sit outside the choice; then either
  // the delete arm or the shared-settings arm — never both. A true delete
  // drops every shared setting (mirrors the per-point early return); an
  // explicit delete:false keeps its element only when no shared setting rides
  // along — alongside settings it would occupy both arms at once.
  const hasShared = DATA_LABELS_SHARED_KEYS.some((key) => opts[key] !== undefined);
  if (opts.delete === true || (opts.delete !== undefined && !hasShared)) {
    return `<c:dLbls>${head}<c:delete val="${opts.delete ? 1 : 0}"/></c:dLbls>`;
  }
  const parts = dataLabelSharedParts(opts, ctx);
  // CT_ChartLines (leaderLines); an empty element is XSD-valid.
  if (opts.showLeaderLines !== undefined)
    parts.push(`<c:showLeaderLines${boolVal(opts.showLeaderLines)}/>`);
  if (opts.leaderLines) parts.push(emptyEl("c:leaderLines"));
  return `<c:dLbls>${head}${parts.join("")}</c:dLbls>`;
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

// ── Plot-area layout XML (CT_Layout > CT_ManualLayout) ──

function stringifyLayout(manualLayout: ManualLayoutOptions | undefined): string {
  if (!manualLayout) return emptyEl("c:layout");
  const parts: string[] = [];
  if (manualLayout.layoutTarget !== undefined)
    parts.push(valEl("c:layoutTarget", manualLayout.layoutTarget));
  if (manualLayout.xMode !== undefined) parts.push(valEl("c:xMode", manualLayout.xMode));
  if (manualLayout.yMode !== undefined) parts.push(valEl("c:yMode", manualLayout.yMode));
  if (manualLayout.wMode !== undefined) parts.push(valEl("c:wMode", manualLayout.wMode));
  if (manualLayout.hMode !== undefined) parts.push(valEl("c:hMode", manualLayout.hMode));
  if (manualLayout.x !== undefined) parts.push(valEl("c:x", manualLayout.x));
  if (manualLayout.y !== undefined) parts.push(valEl("c:y", manualLayout.y));
  if (manualLayout.w !== undefined) parts.push(valEl("c:w", manualLayout.w));
  if (manualLayout.h !== undefined) parts.push(valEl("c:h", manualLayout.h));
  return `<c:layout><c:manualLayout>${parts.join("")}</c:manualLayout></c:layout>`;
}

// ── 3D wall/floor surface XML (CT_Surface) ──

function stringifySurface(
  tag: "c:floor" | "c:sideWall" | "c:backWall",
  opts: SurfaceOptions | undefined,
  ctx: WriteContext,
): string {
  if (!opts) return "";
  const parts: string[] = [];
  // ST_Thickness is a union of a percentage pattern and unsignedInt, but MS
  // Office rejects the "N%" spelling on open — normalize to the bare number.
  if (opts.thickness !== undefined)
    parts.push(
      valEl(
        "c:thickness",
        typeof opts.thickness === "string"
          ? Number(opts.thickness.replace("%", ""))
          : opts.thickness,
      ),
    );
  if (opts.shapeProperties) parts.push(chartSpPr(opts.shapeProperties, ctx));
  return parts.length ? `<${tag}>${parts.join("")}</${tag}>` : `<${tag}/>`;
}

// ── Axis XML (EG_AxShared + CT_CatAx/CT_ValAx/CT_DateAx/CT_SerAx) ──

const AXIS_TAG: Record<AxisKind, string> = {
  category: "c:catAx",
  value: "c:valAx",
  date: "c:dateAx",
  series: "c:serAx",
};

function stringifyScaling(opts: AxisScalingOptions | undefined): string {
  if (!opts) return emptyEl("c:scaling");
  const parts: string[] = [];
  if (opts.logBase !== undefined) parts.push(valEl("c:logBase", opts.logBase));
  if (opts.orientation !== undefined)
    parts.push(valEl("c:orientation", xsdAxisOrientation.to(opts.orientation)));
  if (opts.max !== undefined) parts.push(valEl("c:max", opts.max));
  if (opts.min !== undefined) parts.push(valEl("c:min", opts.min));
  return `<c:scaling>${parts.join("")}</c:scaling>`;
}

function stringifyDisplayUnits(opts: DisplayUnitsOptions): string {
  const parts: string[] = [];
  if (opts.customUnit !== undefined) parts.push(valEl("c:custUnit", opts.customUnit));
  else if (opts.builtInUnit !== undefined) parts.push(valEl("c:builtInUnit", opts.builtInUnit));
  if (opts.label) parts.push(emptyEl("c:dispUnitsLbl"));
  return `<c:dispUnits>${parts.join("")}</c:dispUnits>`;
}

/**
 * Serialize an axis. Field order follows the XSD content model exactly:
 * EG_AxShared (axId…crosses/crossesAt), then the kind-specific tail from
 * CT_CatAx / CT_DateAx / CT_SerAx / CT_ValAx.
 */
function stringifyAxis(opts: WiredAxisOptions, ctx: WriteContext): string {
  const parts: string[] = [];
  parts.push(valEl("c:axId", opts.id));
  parts.push(stringifyScaling(opts.scaling));
  if (opts.delete !== undefined) parts.push(`<c:delete${boolVal(opts.delete)}/>`);
  // axPos is required (EG_AxShared); default by axis role when unset —
  // category-style axes sit at the bottom, value/secondary axes on the left
  const axPos = opts.position ?? (opts.kind === "value" ? "left" : "bottom");
  parts.push(valEl("c:axPos", xsdAxisPosition.to(axPos)));
  parts.push(stringifyChartLines("c:majorGridlines", opts.majorGridlines, ctx));
  parts.push(stringifyChartLines("c:minorGridlines", opts.minorGridlines, ctx));
  if (opts.title !== undefined) {
    // "" is the text-less title placeholder (bare <c:title/>), like the chart title.
    parts.push(
      typeof opts.title === "string" && opts.title === ""
        ? emptyEl("c:title")
        : stringifyTitle(opts.title, ctx),
    );
  }
  if (opts.numberFormat !== undefined) {
    // A plain string keeps the default sourceLinked=1; the object form
    // round-trips an explicit decoupling from the source data.
    const fmt =
      typeof opts.numberFormat === "string" ? { formatCode: opts.numberFormat } : opts.numberFormat;
    const linked = fmt.sourceLinked === undefined || fmt.sourceLinked ? 1 : 0;
    parts.push(`<c:numFmt formatCode="${escapeXml(fmt.formatCode)}" sourceLinked="${linked}"/>`);
  }
  if (opts.majorTickMark !== undefined) parts.push(valEl("c:majorTickMark", opts.majorTickMark));
  if (opts.minorTickMark !== undefined) parts.push(valEl("c:minorTickMark", opts.minorTickMark));
  if (opts.tickLabelPosition !== undefined)
    parts.push(valEl("c:tickLblPos", opts.tickLabelPosition));
  // EG_AxShared: spPr sits between tickLblPos and crossAx; txPr follows spPr.
  parts.push(chartSpPr(opts.shapeProperties, ctx));
  if (opts.textProperties) {
    parts.push(`<c:txPr>${textBodyDesc.stringify(opts.textProperties, ctx) ?? ""}</c:txPr>`);
  }
  parts.push(valEl("c:crossAx", opts.crossAxisId));
  // XSD choice: crosses XOR crossesAt
  if (opts.crossesAt !== undefined) parts.push(valEl("c:crossesAt", opts.crossesAt));
  else if (opts.crosses !== undefined)
    parts.push(valEl("c:crosses", xsdAxisCrosses.to(opts.crosses)));

  switch (opts.kind) {
    case "category":
      if (opts.auto !== undefined) parts.push(`<c:auto${boolVal(opts.auto)}/>`);
      if (opts.labelAlignment !== undefined)
        parts.push(valEl("c:lblAlgn", xsdAxisLabelAlignment.to(opts.labelAlignment)));
      if (opts.labelOffset !== undefined) parts.push(valEl("c:lblOffset", opts.labelOffset));
      if (opts.tickLabelSkip !== undefined) parts.push(valEl("c:tickLblSkip", opts.tickLabelSkip));
      if (opts.tickMarkSkip !== undefined) parts.push(valEl("c:tickMarkSkip", opts.tickMarkSkip));
      if (opts.noMultiLevelLabel !== undefined)
        parts.push(`<c:noMultiLvlLbl${boolVal(opts.noMultiLevelLabel)}/>`);
      break;
    case "date":
      if (opts.auto !== undefined) parts.push(`<c:auto${boolVal(opts.auto)}/>`);
      if (opts.labelOffset !== undefined) parts.push(valEl("c:lblOffset", opts.labelOffset));
      if (opts.baseTimeUnit !== undefined) parts.push(valEl("c:baseTimeUnit", opts.baseTimeUnit));
      if (opts.majorUnit !== undefined) parts.push(valEl("c:majorUnit", opts.majorUnit));
      if (opts.majorTimeUnit !== undefined)
        parts.push(valEl("c:majorTimeUnit", opts.majorTimeUnit));
      if (opts.minorUnit !== undefined) parts.push(valEl("c:minorUnit", opts.minorUnit));
      if (opts.minorTimeUnit !== undefined)
        parts.push(valEl("c:minorTimeUnit", opts.minorTimeUnit));
      break;
    case "series":
      if (opts.tickLabelSkip !== undefined) parts.push(valEl("c:tickLblSkip", opts.tickLabelSkip));
      if (opts.tickMarkSkip !== undefined) parts.push(valEl("c:tickMarkSkip", opts.tickMarkSkip));
      break;
    case "value":
      if (opts.crossBetween !== undefined)
        parts.push(valEl("c:crossBetween", xsdAxisCrossBetween.to(opts.crossBetween)));
      if (opts.majorUnit !== undefined) parts.push(valEl("c:majorUnit", opts.majorUnit));
      if (opts.minorUnit !== undefined) parts.push(valEl("c:minorUnit", opts.minorUnit));
      if (opts.displayUnits) parts.push(stringifyDisplayUnits(opts.displayUnits));
      break;
  }
  const tag = AXIS_TAG[opts.kind];
  return `<${tag}>${parts.join("")}</${tag}>`;
}

// ── Series data XML builders ──

/** Worksheet reference formula (c:f); empty element when not round-tripped. */
function refFormula(formula: string | undefined): string {
  return `<c:f>${escapeXml(formula ?? "")}</c:f>`;
}

function stringifyStrRef(values: readonly (string | undefined)[], formula?: string): string {
  const pts = values
    .map((v, i) => `<c:pt idx="${i}"><c:v>${escapeXml(v ?? "")}</c:v></c:pt>`)
    .join("");
  return `<c:strRef>${refFormula(formula)}<c:strCache><c:ptCount ${attrVal("val", values.length)}/>${pts}</c:strCache></c:strRef>`;
}

function stringifyStrLit(values: readonly string[]): string {
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${escapeXml(v)}</c:v></c:pt>`).join("");
  return `<c:strLit><c:ptCount ${attrVal("val", values.length)}/>${pts}</c:strLit>`;
}

function stringifyMultiLvlStrRef(levels: readonly (readonly string[])[], formula?: string): string {
  const ptCount = levels.reduce((max, lvl) => Math.max(max, lvl.length), 0);
  const lvls = levels
    .map((lvl) => {
      const pts = lvl.map((v, i) => `<c:pt idx="${i}"><c:v>${escapeXml(v)}</c:v></c:pt>`).join("");
      return `<c:lvl>${pts}</c:lvl>`;
    })
    .join("");
  return `<c:multiLvlStrRef>${refFormula(formula)}<c:multiLvlStrCache><c:ptCount ${attrVal("val", ptCount)}/>${lvls}</c:multiLvlStrCache></c:multiLvlStrRef>`;
}

/** Whether the chart carries category-source data (c:cat is optional in CT_Ser). */
function hasCategoryData(opts: ChartSpaceOptions): boolean {
  return (
    opts.categories !== undefined ||
    opts.categoryLabels !== undefined ||
    opts.multiLevelCategories !== undefined ||
    opts.numericCategories === true
  );
}

// CT_AxDataSource choice (c:cat / c:xVal slot): exactly one of multi-level/literal/reference.
function stringifyCategorySource(opts: ChartSpaceOptions): string {
  if (opts.multiLevelCategories)
    return stringifyMultiLvlStrRef(opts.multiLevelCategories, opts.categoryFormula);
  if (opts.categoryLabels) return stringifyStrLit(opts.categoryLabels);
  if (opts.numericCategories) {
    return stringifyNumRef(opts.categories ?? [], opts.categoryFormula, opts.categoryFormatCode);
  }
  return stringifyStrRef(opts.categories ?? [], opts.categoryFormula);
}

// Numeric cache reference for series values and numeric categories (the latter
// keep their literal text so formats survive verbatim).
function stringifyNumRef(
  values: readonly (string | number)[],
  formula?: string,
  formatCode?: string,
): string {
  const pts = values
    .map((v, i) => `<c:pt idx="${i}"><c:v>${typeof v === "number" ? v : escapeXml(v)}</c:v></c:pt>`)
    .join("");
  return `<c:numRef>${refFormula(formula)}<c:numCache><c:formatCode>${escapeXml(formatCode ?? "General")}</c:formatCode><c:ptCount ${attrVal("val", values.length)}/>${pts}</c:numCache></c:numRef>`;
}

// CT_NumData literal form (c:val > c:numLit) — Excel writes this for
// hand-entered series with no worksheet reference.
function stringifyNumLitList(values: readonly number[], formatCode?: string): string {
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("");
  return `<c:numLit><c:formatCode>${escapeXml(formatCode ?? "General")}</c:formatCode><c:ptCount ${attrVal("val", values.length)}/>${pts}</c:numLit>`;
}

// ── Chart type header XML ──

function stringifyUpDownBars(gapWidth: number | undefined): string {
  // CT_UpDownBars: gapWidth? → upBars → downBars (CT_ChartLines, presence-only)
  const inner: string[] = [];
  if (gapWidth !== undefined) inner.push(valEl("c:gapWidth", gapWidth));
  inner.push(emptyEl("c:upBars"));
  inner.push(emptyEl("c:downBars"));
  return `<c:upDownBars>${inner.join("")}</c:upDownBars>`;
}

function stringifyDataTable(opts: DataTableOptions, ctx: WriteContext): string {
  // CT_DTable: showHorzBorder → showVertBorder → showOutline → showKeys → spPr → txPr
  const parts: string[] = [];
  if (opts.showHorizontalBorder !== undefined)
    parts.push(`<c:showHorzBorder${boolVal(opts.showHorizontalBorder)}/>`);
  if (opts.showVerticalBorder !== undefined)
    parts.push(`<c:showVertBorder${boolVal(opts.showVerticalBorder)}/>`);
  if (opts.showOutline !== undefined) parts.push(`<c:showOutline${boolVal(opts.showOutline)}/>`);
  if (opts.showLegendKeys !== undefined) parts.push(`<c:showKeys${boolVal(opts.showLegendKeys)}/>`);
  if (opts.shapeProperties) parts.push(chartSpPr(opts.shapeProperties, ctx));
  if (opts.textProperties) {
    parts.push(`<c:txPr>${textBodyDesc.stringify(opts.textProperties, ctx) ?? ""}</c:txPr>`);
  }
  return `<c:dTable>${parts.join("")}</c:dTable>`;
}

function stringifyProtection(opts: ProtectionOptions): string {
  const parts: string[] = [];
  if (opts.chartObject !== undefined) parts.push(`<c:chartObject${boolVal(opts.chartObject)}/>`);
  if (opts.data !== undefined) parts.push(`<c:data${boolVal(opts.data)}/>`);
  if (opts.formatting !== undefined) parts.push(`<c:formatting${boolVal(opts.formatting)}/>`);
  if (opts.selection !== undefined) parts.push(`<c:selection${boolVal(opts.selection)}/>`);
  if (opts.userInterface !== undefined)
    parts.push(`<c:userInterface${boolVal(opts.userInterface)}/>`);
  return `<c:protection>${parts.join("")}</c:protection>`;
}

function stringifyExternalData(opts: ExternalDataOptions): string | undefined {
  // r:id must resolve to a relationship the chart part owns — an unresolvable
  // id (fresh authoring, junk reference) would dangle, so skip the element
  if (!opts.relationshipId || /^r?Id?(0|undefined)?$/i.test(opts.relationshipId)) return undefined;
  const auto =
    opts.autoUpdate !== undefined ? `<c:autoUpdate val="${opts.autoUpdate ? 1 : 0}"/>` : "";
  return `<c:externalData r:id="${escapeXml(opts.relationshipId)}">${auto}</c:externalData>`;
}

function stringifyHeaderFooter(opts: ChartHeaderFooterOptions): string {
  const parts: string[] = [];
  if (opts.oddHeader !== undefined)
    parts.push(`<c:oddHeader>${escapeXml(opts.oddHeader)}</c:oddHeader>`);
  if (opts.oddFooter !== undefined)
    parts.push(`<c:oddFooter>${escapeXml(opts.oddFooter)}</c:oddFooter>`);
  if (opts.evenHeader !== undefined)
    parts.push(`<c:evenHeader>${escapeXml(opts.evenHeader)}</c:evenHeader>`);
  if (opts.evenFooter !== undefined)
    parts.push(`<c:evenFooter>${escapeXml(opts.evenFooter)}</c:evenFooter>`);
  if (opts.firstHeader !== undefined)
    parts.push(`<c:firstHeader>${escapeXml(opts.firstHeader)}</c:firstHeader>`);
  if (opts.firstFooter !== undefined)
    parts.push(`<c:firstFooter>${escapeXml(opts.firstFooter)}</c:firstFooter>`);
  const attrs: string[] = [];
  if (opts.alignWithMargins !== undefined)
    attrs.push(`alignWithMargins="${opts.alignWithMargins ? 1 : 0}"`);
  if (opts.differentOddEven !== undefined)
    attrs.push(`differentOddEven="${opts.differentOddEven ? 1 : 0}"`);
  if (opts.differentFirst !== undefined)
    attrs.push(`differentFirst="${opts.differentFirst ? 1 : 0}"`);
  const attrStr = attrs.length ? " " + attrs.join(" ") : "";
  if (parts.length === 0) return `<c:headerFooter${attrStr}/>`;
  return `<c:headerFooter${attrStr}>${parts.join("")}</c:headerFooter>`;
}

function stringifyPageMargins(opts: ChartPageMarginsOptions): string {
  // CT_PageMargins attributes are all XSD use="required". Fill Office default
  // margins (inches) for any the caller omits so emission is always schema-valid.
  const l = opts.left ?? 0.7;
  const r = opts.right ?? 0.7;
  const t = opts.top ?? 0.75;
  const b = opts.bottom ?? 0.75;
  const header = opts.header ?? 0.3;
  const footer = opts.footer ?? 0.3;
  return `<c:pageMargins l="${l}" r="${r}" t="${t}" b="${b}" header="${header}" footer="${footer}"/>`;
}

function stringifyPageSetup(opts: ChartPageSetupOptions): string {
  const attrs: string[] = [];
  if (opts.paperSize !== undefined) attrs.push(`paperSize="${opts.paperSize}"`);
  if (opts.paperHeight !== undefined) attrs.push(`paperHeight="${escapeXml(opts.paperHeight)}"`);
  if (opts.paperWidth !== undefined) attrs.push(`paperWidth="${escapeXml(opts.paperWidth)}"`);
  if (opts.firstPageNumber !== undefined) attrs.push(`firstPageNumber="${opts.firstPageNumber}"`);
  if (opts.orientation !== undefined) attrs.push(`orientation="${opts.orientation}"`);
  if (opts.blackAndWhite !== undefined) attrs.push(`blackAndWhite="${opts.blackAndWhite ? 1 : 0}"`);
  if (opts.draft !== undefined) attrs.push(`draft="${opts.draft ? 1 : 0}"`);
  if (opts.useFirstPageNumber !== undefined)
    attrs.push(`useFirstPageNumber="${opts.useFirstPageNumber ? 1 : 0}"`);
  if (opts.horizontalDpi !== undefined) attrs.push(`horizontalDpi="${opts.horizontalDpi}"`);
  if (opts.verticalDpi !== undefined) attrs.push(`verticalDpi="${opts.verticalDpi}"`);
  if (opts.copies !== undefined) attrs.push(`copies="${opts.copies}"`);
  if (attrs.length === 0) return `<c:pageSetup/>`;
  return `<c:pageSetup ${attrs.join(" ")}/>`;
}

function stringifyPrintSettings(opts: PrintSettingsOptions): string {
  const parts: string[] = [];
  if (opts.headerFooter) parts.push(stringifyHeaderFooter(opts.headerFooter));
  if (opts.pageMargins) parts.push(stringifyPageMargins(opts.pageMargins));
  if (opts.pageSetup) parts.push(stringifyPageSetup(opts.pageSetup));
  if (opts.legacyDrawingId !== undefined)
    parts.push(`<c:legacyDrawingHF r:id="${escapeXml(opts.legacyDrawingId)}"/>`);
  return `<c:printSettings>${parts.join("")}</c:printSettings>`;
}

// Header fields shared by the main group and combo secondary groups.
interface ChartGroupHeaderFields {
  type: ChartType;
  grouping?: ChartGrouping;
  radarStyle?: RadarStyle;
  ofPieType?: OfPieType;
  wireframe?: boolean;
  varyColors?: boolean;
}

function chartTypeHeader(opts: ChartGroupHeaderFields & { threeD?: boolean }): string {
  const tag = opts.threeD ? CHART_TYPE_TAGS_3D[opts.type] : CHART_TYPE_TAGS[opts.type];
  if (!tag) throw new Error(`Unsupported chart type: ${opts.type}`);

  const headerParts: string[] = [];

  switch (opts.type) {
    case "column":
    case "bar":
      headerParts.push(valEl("c:barDir", opts.type === "column" ? "col" : "bar"));
      headerParts.push(valEl("c:grouping", opts.grouping ?? "clustered"));
      break;
    case "line":
    case "area":
      headerParts.push(valEl("c:grouping", opts.grouping ?? "standard"));
      break;
    case "scatter":
      headerParts.push(valEl("c:scatterStyle", "line"));
      break;
    case "radar":
      headerParts.push(valEl("c:radarStyle", opts.radarStyle ?? "standard"));
      break;
    case "ofPie":
      headerParts.push(valEl("c:ofPieType", opts.ofPieType ?? "pie"));
      break;
    case "surface":
      // EG_SurfaceChartShared head: wireframe precedes ser.
      if (opts.wireframe !== undefined)
        headerParts.push(`<c:wireframe${boolVal(opts.wireframe)}/>`);
      break;
  }

  // c:varyColors sits right before ser in every chart-group content model;
  // emitted only when the source carried it (corpus is mixed).
  if (opts.varyColors !== undefined) headerParts.push(`<c:varyColors${boolVal(opts.varyColors)}/>`);

  return `<${tag}>${headerParts.join("")}`;
}

function chartTypeFooter(opts: ChartSpaceOptions, ctx: WriteContext): string {
  const tag = opts.threeD ? CHART_TYPE_TAGS_3D[opts.type] : CHART_TYPE_TAGS[opts.type];
  const parts: string[] = [];

  // Type-specific children between ser and axId, in XSD document order.
  // Group-tail elements (dropLines for line/area) lead because they sit inside
  // EG_*ChartShared right after ser/dLbls, before the outer-type decorations.
  switch (opts.type) {
    case "column":
    case "bar":
      if (opts.gapWidth !== undefined) parts.push(valEl("c:gapWidth", opts.gapWidth));
      if (opts.threeD) {
        // CT_Bar3DChart: gapWidth → gapDepth → shape
        if (opts.gapDepth !== undefined) parts.push(valEl("c:gapDepth", opts.gapDepth));
        if (opts.shape !== undefined) parts.push(valEl("c:shape", opts.shape));
      } else {
        // CT_BarChart: gapWidth → overlap → serLines
        if (opts.overlap !== undefined) parts.push(valEl("c:overlap", opts.overlap));
        parts.push(stringifyChartLines("c:serLines", opts.seriesLines, ctx));
      }
      break;
    case "line":
      // EG_LineChartShared tail: dropLines (after ser/dLbls)
      parts.push(stringifyChartLines("c:dropLines", opts.dropLines, ctx));
      if (!opts.threeD) {
        // CT_LineChart: dropLines → hiLowLines → upDownBars → marker → smooth
        parts.push(stringifyChartLines("c:hiLowLines", opts.highLowLines, ctx));
        if (opts.upDownBars) parts.push(stringifyUpDownBars(opts.upDownBarsGapWidth));
        if (opts.markers !== undefined) parts.push(`<c:marker${boolVal(opts.markers)}/>`);
        if (opts.smooth !== undefined) parts.push(`<c:smooth${boolVal(opts.smooth)}/>`);
      } else if (opts.gapDepth !== undefined) {
        // CT_Line3DChart: dropLines → gapDepth
        parts.push(valEl("c:gapDepth", opts.gapDepth));
      }
      break;
    case "area":
      // EG_AreaChartShared tail: dropLines
      parts.push(stringifyChartLines("c:dropLines", opts.dropLines, ctx));
      if (opts.threeD && opts.gapDepth !== undefined)
        parts.push(valEl("c:gapDepth", opts.gapDepth));
      break;
    case "stock":
      // CT_StockChart: dropLines → hiLowLines → upDownBars
      parts.push(stringifyChartLines("c:dropLines", opts.dropLines, ctx));
      parts.push(stringifyChartLines("c:hiLowLines", opts.highLowLines, ctx));
      if (opts.upDownBars) parts.push(stringifyUpDownBars(opts.upDownBarsGapWidth));
      break;
    case "pie":
      // CT_PieChart: firstSliceAng (absent on CT_Pie3DChart).
      if (!opts.threeD && opts.firstSliceAngle !== undefined)
        parts.push(valEl("c:firstSliceAng", opts.firstSliceAngle));
      break;
    case "doughnut":
      // CT_DoughnutChart: firstSliceAng → holeSize. holeSize is XSD-optional
      // (default 10) but the Open XML SDK's stricter schema requires it —
      // always emit so both validators pass with identical rendering.
      if (opts.firstSliceAngle !== undefined)
        parts.push(valEl("c:firstSliceAng", opts.firstSliceAngle));
      parts.push(valEl("c:holeSize", opts.holeSize ?? 10));
      break;
    case "bubble":
      // CT_BubbleChart: bubbleScale → showNegBubbles → sizeRepresents
      if (opts.bubbleScale !== undefined) parts.push(valEl("c:bubbleScale", opts.bubbleScale));
      if (opts.showNegativeBubbles !== undefined)
        parts.push(`<c:showNegBubbles${boolVal(opts.showNegativeBubbles)}/>`);
      if (opts.sizeRepresents !== undefined)
        parts.push(valEl("c:sizeRepresents", xsdSizeRepresents.to(opts.sizeRepresents)));
      break;
    case "ofPie":
      // CT_OfPieChart: gapWidth → splitType → splitPos → custSplit → secondPieSize → serLines
      if (opts.gapWidth !== undefined) parts.push(valEl("c:gapWidth", opts.gapWidth));
      if (opts.splitType !== undefined)
        parts.push(valEl("c:splitType", xsdSplitType.to(opts.splitType)));
      if (opts.splitPosition !== undefined) parts.push(valEl("c:splitPos", opts.splitPosition));
      if (opts.customSplitPoints?.length) {
        const pts = opts.customSplitPoints.map((p) => valEl("c:secondPiePt", p)).join("");
        parts.push(`<c:custSplit>${pts}</c:custSplit>`);
      }
      if (opts.secondPieSize !== undefined)
        parts.push(valEl("c:secondPieSize", opts.secondPieSize));
      parts.push(stringifyChartLines("c:serLines", opts.seriesLines, ctx));
      break;
    case "surface":
      // EG_SurfaceChartShared tail: bandFmts (after ser, before axId)
      if (opts.bandFormats?.length) parts.push(stringifyBandFormats(opts.bandFormats));
      break;
  }

  // Axes (pie/doughnut have none). The axId sequence must reference the axes
  // actually emitted below — use the round-tripped sequence when present
  // (legacy files may carry dangling axId val="0"), else derive it from the
  // effective axis list, never hardcode. axisIds only applies alongside an
  // explicit axes list; without axes the default axes carry their own ids.
  if (!NO_AXES_TYPES.has(opts.type)) {
    const ids = opts.axes && opts.axisIds ? opts.axisIds : axesFor(opts).map((axis) => axis.id);
    for (const id of ids) parts.push(valEl("c:axId", id));
  }

  return `${parts.join("")}</${tag}>`;
}

// ── Series sub-elements (CT_Marker / CT_DPt / CT_PictureOptions) ──

function stringifyMarker(opts: MarkerOptions, ctx: WriteContext): string {
  const parts: string[] = [];
  if (opts.symbol !== undefined) parts.push(valEl("c:symbol", opts.symbol));
  if (opts.size !== undefined) parts.push(valEl("c:size", opts.size));
  parts.push(chartSpPr(opts.shapeProperties, ctx));
  return `<c:marker>${parts.join("")}</c:marker>`;
}

function stringifyPictureOptions(opts: PictureOptionsOptions): string {
  const parts: string[] = [];
  if (opts.applyToFront !== undefined) parts.push(`<c:applyToFront${boolVal(opts.applyToFront)}/>`);
  if (opts.applyToSides !== undefined) parts.push(`<c:applyToSides${boolVal(opts.applyToSides)}/>`);
  if (opts.applyToEnd !== undefined) parts.push(`<c:applyToEnd${boolVal(opts.applyToEnd)}/>`);
  if (opts.pictureFormat !== undefined) parts.push(valEl("c:pictureFormat", opts.pictureFormat));
  if (opts.pictureStackUnit !== undefined)
    parts.push(valEl("c:pictureStackUnit", opts.pictureStackUnit));
  return `<c:pictureOptions>${parts.join("")}</c:pictureOptions>`;
}

function stringifyDataPoint(opts: DataPointOptions, ctx: WriteContext): string {
  // CT_DPt: idx → invertIfNegative → marker → bubble3D → explosion → spPr → pictureOptions
  const parts: string[] = [valEl("c:idx", opts.index)];
  if (opts.invertIfNegative !== undefined)
    parts.push(`<c:invertIfNegative${boolVal(opts.invertIfNegative)}/>`);
  if (opts.marker) parts.push(stringifyMarker(opts.marker, ctx));
  if (opts.bubble3D !== undefined) parts.push(`<c:bubble3D${boolVal(opts.bubble3D)}/>`);
  if (opts.explosion !== undefined) parts.push(valEl("c:explosion", opts.explosion));
  parts.push(chartSpPr(opts.shapeProperties, ctx));
  if (opts.pictureOptions) parts.push(stringifyPictureOptions(opts.pictureOptions));
  return `<c:dPt>${parts.join("")}</c:dPt>`;
}

// ── Series XML ──

function stringifySeries(
  index: number,
  series: ChartSeriesData | ScatterSeriesData,
  opts: ChartSpaceOptions,
  ctx: WriteContext,
  groupType?: ChartType,
): string {
  // A combo chart's second group keeps its own type for the per-series
  // branches below (marker on line, invertIfNegative on bar, …).
  const chartType = groupType ?? opts.type;
  const parts: string[] = [];
  const s = series as ChartSeriesData;

  // EG_SerShared: idx, order, tx, spPr. Round-tripped series keep their
  // chart-wide idx/order (combo groups interleave them); fresh series number
  // by position.
  parts.push(valEl("c:idx", s.index ?? index));
  parts.push(valEl("c:order", s.order ?? index));
  if (series.nameLiteral && series.name !== undefined) {
    parts.push(`<c:tx><c:v>${escapeXml(series.name)}</c:v></c:tx>`);
  } else if (series.name !== undefined || series.nameFormula !== undefined)
    parts.push(`<c:tx>${stringifyStrRef([series.name], series.nameFormula)}</c:tx>`);
  // Series spPr is optional in CT_Ser — emit only when the source carried one
  // (round-trip); fresh series keep the bare form like legacy Word files.
  parts.push(chartSpPr(s.shapeProperties, ctx));

  // type-specific head before dPt (per CT_xxxSer content model)
  if (chartType === "line" || chartType === "scatter" || chartType === "radar") {
    if (s.marker) parts.push(stringifyMarker(s.marker, ctx));
  } else if (chartType === "column" || chartType === "bar") {
    if (s.invertIfNegative !== undefined)
      parts.push(`<c:invertIfNegative${boolVal(s.invertIfNegative)}/>`);
    if (s.pictureOptions) parts.push(stringifyPictureOptions(s.pictureOptions));
  } else if (chartType === "area") {
    if (s.pictureOptions) parts.push(stringifyPictureOptions(s.pictureOptions));
  } else if (chartType === "bubble") {
    if (s.invertIfNegative !== undefined)
      parts.push(`<c:invertIfNegative${boolVal(s.invertIfNegative)}/>`);
  } else if (chartType === "pie" || chartType === "doughnut" || chartType === "ofPie") {
    if (s.explosion !== undefined) parts.push(valEl("c:explosion", s.explosion));
  }

  // shared decorations (XSD order: dPt → dLbls → trendline → errBars)
  if (s.dataPoints) {
    for (const dp of s.dataPoints) parts.push(stringifyDataPoint(dp, ctx));
  }
  if (s.dataLabels) parts.push(stringifyDataLabels(s.dataLabels, ctx));
  if (s.trendlines) {
    for (const tl of s.trendlines) parts.push(stringifyTrendline(tl));
  }
  if (s.errorBars) parts.push(stringifyErrBars(s.errorBars));

  // data references (type-specific element names)
  if (chartType === "bubble") {
    const bs = series as BubbleSeriesData;
    parts.push(`<c:xVal>${stringifyNumRef(bs.xValues)}</c:xVal>`);
    parts.push(`<c:yVal>${stringifyNumRef(bs.yValues)}</c:yVal>`);
    parts.push(`<c:bubbleSize>${stringifyNumRef(bs.bubbleSize)}</c:bubbleSize>`);
  } else if (chartType === "scatter") {
    if ("xValues" in series) {
      // True numeric axes: c:xVal/c:yVal are CT_NumDataSource references.
      parts.push(`<c:xVal>${stringifyNumRef(series.xValues)}</c:xVal>`);
      parts.push(
        `<c:yVal>${stringifyNumRef(series.yValues, series.valueFormula, series.formatCode)}</c:yVal>`,
      );
    } else {
      // Label-x shape: x values share the category source model — string
      // labels round-trip through c:strRef like regular categories.
      parts.push(`<c:xVal>${stringifyCategorySource(opts)}</c:xVal>`);
      parts.push(`<c:yVal>${stringifyNumRef(s.values, s.valueFormula, s.formatCode)}</c:yVal>`);
    }
  } else {
    if (hasCategoryData(opts)) {
      parts.push(`<c:cat>${stringifyCategorySource(opts)}</c:cat>`);
    }
    parts.push(
      `<c:val>${s.valueLiteral ? stringifyNumLitList(s.values, s.formatCode) : stringifyNumRef(s.values, s.valueFormula, s.formatCode)}</c:val>`,
    );
  }

  // type-specific tail after data references
  if ((chartType === "line" || chartType === "scatter") && s.smooth !== undefined) {
    parts.push(`<c:smooth${boolVal(s.smooth)}/>`);
  }
  if ((chartType === "column" || chartType === "bar") && s.shape !== undefined) {
    parts.push(valEl("c:shape", s.shape));
  }
  if (chartType === "bubble" && s.bubble3D !== undefined) {
    parts.push(`<c:bubble3D${boolVal(s.bubble3D)}/>`);
  }

  // CT_xxxSer tail — verbatim series extLst (c16:uniqueId's home).
  if (s.ext) parts.push(`<c:extLst>${s.ext}</c:extLst>`);

  return `<c:ser>${parts.join("")}</c:ser>`;
}

// CT_ChartLines: a bare element or an spPr carrier (true keeps the bare form)
function stringifyChartLines(
  tag: string,
  opts: boolean | ChartLinesOptions | undefined,
  ctx: WriteContext,
): string {
  if (opts === undefined || opts === false) return "";
  if (opts === true) return emptyEl(tag);
  return `<${tag}>${chartSpPr(opts.shapeProperties, ctx)}</${tag}>`;
}

/** Read any CT_ChartLines element: true when bare, decorations when carried. */
function readChartLines(
  parent: XmlElement,
  tag: string,
  ctx: ReadContext,
): boolean | ChartLinesOptions | undefined {
  const el = findChild(parent, tag);
  if (!el) return undefined;
  const spPrEl = findChild(el, "c:spPr");
  if (!spPrEl) return true;
  const lines: ChartLinesOptions = {};
  const shapeProperties = shapePropertiesDesc.parse(spPrEl, ctx);
  if (shapeProperties) lines.shapeProperties = shapeProperties;
  return lines;
}

/** Read one secondary chart group (combo): own series, flags, and 2D group tail. */
function readSecondaryGroup(
  secondaryEl: XmlElement,
  ctx: ReadContext,
): SecondaryChartGroupOptions | undefined {
  const mapping = TAG_TO_CHART_TYPE[secondaryEl.name ?? ""];
  if (!mapping) return undefined;
  let secType: ChartType = mapping.type;
  const isBarTag = secondaryEl.name === "c:barChart" || secondaryEl.name === "c:bar3DChart";
  if (isBarTag) {
    const barDir = findChild(secondaryEl, "c:barDir");
    const barDirVal = barDir ? attr(barDir, "val") : undefined;
    if (barDirVal === "bar") secType = "bar";
    else if (barDirVal === "col") secType = "column";
  }
  const g: SecondaryChartGroupOptions = { type: secType, series: [] };
  const grouping = readValStr(secondaryEl, "c:grouping");
  if (grouping) g.grouping = grouping as ChartGrouping;
  const varyColors = readBoolAttr(secondaryEl, "c:varyColors");
  if (varyColors !== undefined) g.varyColors = varyColors;
  const groupLabels = readDataLabels(secondaryEl, ctx);
  if (groupLabels) g.dataLabels = groupLabels;
  // Group tails in XSD document order — same elements the main group reads.
  if (secType === "bar" || secType === "column") {
    const gapWidth = readValNum(secondaryEl, "c:gapWidth");
    if (gapWidth !== undefined) g.gapWidth = gapWidth;
    if (!mapping.threeD) {
      const overlap = readValNum(secondaryEl, "c:overlap");
      if (overlap !== undefined) g.overlap = overlap;
    } else {
      const gapDepth = readValNum(secondaryEl, "c:gapDepth");
      if (gapDepth !== undefined) g.gapDepth = gapDepth;
    }
    const seriesLines = readChartLines(secondaryEl, "c:serLines", ctx);
    if (seriesLines !== undefined) g.seriesLines = seriesLines;
  } else if (secType === "line") {
    const dropLines = readChartLines(secondaryEl, "c:dropLines", ctx);
    if (dropLines !== undefined) g.dropLines = dropLines;
    const highLowLines = readChartLines(secondaryEl, "c:hiLowLines", ctx);
    if (highLowLines !== undefined) g.highLowLines = highLowLines;
    const upDownBars = findChild(secondaryEl, "c:upDownBars");
    if (upDownBars) {
      g.upDownBars = true;
      const gw = readValNum(upDownBars, "c:gapWidth");
      if (gw !== undefined) g.upDownBarsGapWidth = gw;
    }
    const markers = readBoolAttr(secondaryEl, "c:marker");
    if (markers !== undefined) g.markers = markers;
    const smooth = readBoolAttr(secondaryEl, "c:smooth");
    if (smooth !== undefined) g.smooth = smooth;
  } else if (secType === "area") {
    const dropLines = readChartLines(secondaryEl, "c:dropLines", ctx);
    if (dropLines !== undefined) g.dropLines = dropLines;
  }
  const secAxIds = children(secondaryEl, "c:axId");
  if (secAxIds.length > 0) {
    const ids = secAxIds.map((e) => Number(attr(e, "val"))).filter((n) => !isNaN(n));
    if (ids.length > 0) g.axisIds = ids;
  }
  const secSeries: ChartSeriesData[] = [];
  for (const serEl of children(secondaryEl, "c:ser")) {
    const { name, literal: nameLiteral } = readSeriesName(serEl);
    const txEl = findChild(serEl, "c:tx");
    const nameFormula = txEl ? readRefMeta(txEl).formula : undefined;
    const valueEl = findChild(serEl, "c:val") ?? findChild(serEl, "c:yVal");
    const valueLiteral = valueEl ? findChild(valueEl, "c:numLit") !== undefined : false;
    const values = valueEl ? readNumCache(valueEl) : [];
    const valMeta = valueEl ? readRefMeta(valueEl) : {};
    secSeries.push({
      name,
      ...(nameLiteral ? { nameLiteral } : {}),
      ...(nameFormula ? { nameFormula } : {}),
      ...(valueLiteral ? { valueLiteral } : {}),
      ...(valMeta.formula ? { valueFormula: valMeta.formula } : {}),
      ...(valMeta.formatCode ? { formatCode: valMeta.formatCode } : {}),
      values,
      ...readSeriesCommon(serEl, ctx),
    });
  }
  g.series = secSeries;
  return g;
}

// ── Secondary chart group (combo charts) ──

function stringifySecondaryGroup(
  g: SecondaryChartGroupOptions,
  opts: ChartSpaceOptions,
  ctx: WriteContext,
): string {
  const parts: string[] = [chartTypeHeader(g)];
  for (const [i, series] of (g.series ?? []).entries()) {
    parts.push(stringifySeries(i, series, opts, ctx, g.type));
  }
  if (g.dataLabels) parts.push(stringifyDataLabels(g.dataLabels, ctx));
  // Group tails in XSD document order (2D secondary groups only — same order
  // chartTypeFooter emits for the main group).
  switch (g.type) {
    case "bar":
    case "column":
      // CT_BarChart: gapWidth → overlap → serLines
      if (g.gapWidth !== undefined) parts.push(valEl("c:gapWidth", g.gapWidth));
      if (g.overlap !== undefined) parts.push(valEl("c:overlap", g.overlap));
      parts.push(stringifyChartLines("c:serLines", g.seriesLines, ctx));
      break;
    case "line":
      // CT_LineChart: dropLines → hiLowLines → upDownBars → marker → smooth
      parts.push(stringifyChartLines("c:dropLines", g.dropLines, ctx));
      parts.push(stringifyChartLines("c:hiLowLines", g.highLowLines, ctx));
      if (g.upDownBars) parts.push(stringifyUpDownBars(g.upDownBarsGapWidth));
      if (g.markers !== undefined) parts.push(`<c:marker${boolVal(g.markers)}/>`);
      if (g.smooth !== undefined) parts.push(`<c:smooth${boolVal(g.smooth)}/>`);
      break;
    case "area":
      // EG_AreaChartShared tail: dropLines
      parts.push(stringifyChartLines("c:dropLines", g.dropLines, ctx));
      break;
    case "scatter":
      // CT_ScatterChart tail is nothing beyond axId.
      break;
  }
  for (const id of g.axisIds ?? []) parts.push(valEl("c:axId", id));
  // Secondary groups are 2D-only; the header emitted the opening tag, so the
  // group element must close here or the chart part is not well-formed.
  return `${parts.join("")}</${CHART_TYPE_TAGS[g.type]}>`;
}

// ── Title XML ──

// Same c:rich run shape for chart and axis titles so parse reuses readTitleText.
function stringifyTitle(title: string | ChartTitleOptions, ctx: WriteContext): string {
  // A bare string is a text-only title; every other field stays unset.
  const opts: ChartTitleOptions = typeof title === "string" ? { text: title } : title;
  const parts: string[] = [];
  if (opts.text !== undefined) {
    parts.push(
      typeof opts.text === "string"
        ? `<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeXml(opts.text)}</a:t></a:r></a:p></c:rich></c:tx>`
        : `<c:tx><c:rich>${textBodyDesc.stringify(opts.text, ctx) ?? ""}</c:rich></c:tx>`,
    );
  }
  if (opts.layout !== undefined) {
    parts.push(
      opts.layout === true ? "<c:layout/>" : stringifyLayout(opts.layout as ManualLayoutOptions),
    );
  }
  if (opts.overlay !== undefined) parts.push(`<c:overlay${boolVal(opts.overlay)}/>`);
  if (opts.shapeProperties) parts.push(chartSpPr(opts.shapeProperties, ctx));
  if (opts.textProperties) {
    parts.push(`<c:txPr>${textBodyDesc.stringify(opts.textProperties, ctx) ?? ""}</c:txPr>`);
  }
  return `<c:title>${parts.join("")}</c:title>`;
}

// ── Legend XML ──

function stringifyLegendEntry(entry: LegendEntryOptions, ctx: WriteContext): string {
  // CT_LegendEntry: idx → choice(delete | EG_LegendEntryData[txPr]).
  if (entry.delete) {
    return `<c:legendEntry><c:idx val="${entry.index}"/><c:delete val="1"/></c:legendEntry>`;
  }
  if (entry.textProperties) {
    const txPr = textBodyDesc.stringify(entry.textProperties, ctx) ?? "";
    return `<c:legendEntry><c:idx val="${entry.index}"/><c:txPr>${txPr}</c:txPr></c:legendEntry>`;
  }
  return "";
}

function stringifyLegend(opts: ChartSpaceOptions, ctx: WriteContext): string {
  // A parsed legend that carried no c:legendPos must stay without one
  // (showLegend===true marks the round-trip path); fresh charts keep the
  // default bottom-position emission.
  const pos =
    opts.legendPosition !== undefined
      ? `<c:legendPos val="${xsdLegendPosition.to(opts.legendPosition)}"/>`
      : opts.showLegend === true
        ? ""
        : `<c:legendPos val="b"/>`;
  const entries = (opts.legendEntries ?? []).map((e) => stringifyLegendEntry(e, ctx)).join("");
  const layout =
    opts.legendLayout === true
      ? "<c:layout/>"
      : opts.legendLayout
        ? stringifyLayout(opts.legendLayout)
        : "";
  const overlay =
    opts.legendOverlay !== undefined ? `<c:overlay${boolVal(opts.legendOverlay)}/>` : "";
  const spPr = chartSpPr(opts.legendShapeProperties, ctx);
  const txPr = opts.legendTextProperties
    ? `<c:txPr>${textBodyDesc.stringify(opts.legendTextProperties, ctx) ?? ""}</c:txPr>`
    : "";
  return `<c:legend>${pos}${entries}${layout}${overlay}${spPr}${txPr}</c:legend>`;
}

function stringifyPivotSource(opts: PivotSourceOptions): string {
  return `<c:pivotSource><c:name>${escapeXml(opts.name)}</c:name><c:fmtId val="${opts.formatId}"/></c:pivotSource>`;
}

function stringifyPivotFormats(
  formats: readonly ChartPivotFormatOptions[],
  ctx: WriteContext,
): string {
  const inner = formats
    .map((f) => {
      let entry = `<c:idx val="${f.index}"/>`;
      if (f.marker) entry += stringifyMarker(f.marker, ctx);
      return `<c:pivotFmt>${entry}</c:pivotFmt>`;
    })
    .join("");
  return `<c:pivotFmts>${inner}</c:pivotFmts>`;
}

function stringifyBandFormats(formats: readonly BandFormatOptions[]): string {
  const inner = formats.map((f) => `<c:bandFmt><c:idx val="${f.index}"/></c:bandFmt>`).join("");
  return `<c:bandFmts>${inner}</c:bandFmts>`;
}

// ── Shared boilerplate ──

/**
 * Wrap c:spPr inner content. An empty shape-properties object is a real
 * `<c:spPr/>` placeholder (presence itself is the content), so undefined
 * inner still emits the bare element.
 */
function chartSpPr(opts: ShapePropertiesOptions | undefined, ctx: WriteContext): string {
  if (!opts) return "";
  const inner = shapePropertiesDesc.stringify(opts, ctx);
  return `<c:spPr>${inner ?? ""}</c:spPr>`;
}

// ── Axes XML ──

/** AxisOptions with the internal wiring ids guaranteed present. */
type WiredAxisOptions = AxisOptions & { id: number; crossAxisId: number };

/** Default bottom category axis (c:catAx) for generated charts. */
function categoryAxis(id: number, crossAxisId: number): WiredAxisOptions {
  return {
    kind: "category",
    id,
    crossAxisId,
    scaling: { orientation: "ascending" },
    delete: false,
    position: "bottom",
    crosses: "zero",
    auto: true,
    labelOffset: 100,
    noMultiLevelLabel: false,
  };
}

/** Default left value axis (c:valAx). */
function valueAxis(id: number, crossAxisId: number): WiredAxisOptions {
  return {
    kind: "value",
    id,
    crossAxisId,
    scaling: { orientation: "ascending" },
    delete: false,
    position: "left",
    numberFormat: "General",
    crosses: "zero",
  };
}

/** Default series axis (c:serAx) for surface charts. */
function seriesAxis(id: number, crossAxisId: number): WiredAxisOptions {
  return {
    kind: "series",
    id,
    crossAxisId,
    scaling: { orientation: "ascending" },
    delete: false,
    position: "bottom",
    numberFormat: "General",
    crosses: "zero",
    tickLabelSkip: 1,
  };
}

/** Sensible default axes derived from chart type, matching prior hardcoded output. */
function defaultAxesFor(chartType: ChartType, threeD?: boolean): readonly WiredAxisOptions[] {
  if (NO_AXES_TYPES.has(chartType)) return [];
  if (chartType === "scatter" || chartType === "bubble") {
    return [valueAxis(10, 20), valueAxis(20, 10)];
  }
  if (chartType === "surface") {
    return [categoryAxis(10, 20), valueAxis(20, 10), seriesAxis(30, 10)];
  }
  if (chartType === "stock") {
    return [categoryAxis(10, 20), valueAxis(20, 10)];
  }
  if (threeD) {
    // 3D bar/line groups take cat + val + series axes (the third axis of a
    // 3-axis group is c:serAx — emitting a second c:catAx breaks MS Office).
    return [categoryAxis(10, 20), valueAxis(20, 10), seriesAxis(30, 10)];
  }
  return [categoryAxis(10, 20), valueAxis(20, 10)];
}

/** Provided axes override defaults; otherwise defaults are derived from chart type. */
function axesFor(opts: ChartSpaceOptions): readonly WiredAxisOptions[] {
  if (!opts.axes) return defaultAxesFor(opts.type, opts.threeD);
  // Fresh axes may omit their ids (nothing outside the axis pair references
  // them) — fill from the same id slots the default factories use so a
  // partial customization keeps consistent cat/val/ser wiring.
  const slots = defaultAxesFor(opts.type, opts.threeD);
  return opts.axes.map((axis, i) => {
    const slot = slots[i] ?? slots[slots.length - 1]!;
    return {
      ...axis,
      id: axis.id ?? slot.id,
      crossAxisId: axis.crossAxisId ?? slot.crossAxisId,
    };
  });
}

function stringifyAxes(opts: ChartSpaceOptions, ctx: WriteContext): string {
  return axesFor(opts)
    .map((axis) => stringifyAxis(axis, ctx))
    .join("");
}

// ── Read helpers ──

function readStrCache(el: XmlElement): string[] {
  const strRef = findChild(el, "c:strRef");
  if (!strRef) return [];
  const strCache = findChild(strRef, "c:strCache");
  if (!strCache?.elements) return [];
  const result: string[] = [];
  for (const pt of strCache.elements) {
    if (pt.name === "c:pt") {
      const v = findChild(pt, "c:v");
      const text = v ? textOf(v) : "";
      if (text !== "") result.push(text);
    }
  }
  return result;
}

/** Numeric category labels read as literal text (c:cat > c:numRef > c:numCache). */
function readNumCacheText(el: XmlElement): string[] {
  const numRef = findChild(el, "c:numRef");
  if (!numRef) return [];
  const numCache = findChild(numRef, "c:numCache");
  if (!numCache?.elements) return [];
  const result: string[] = [];
  for (const pt of numCache.elements) {
    if (pt.name === "c:pt") {
      const v = findChild(pt, "c:v");
      const text = v ? (textOf(v) ?? "") : "";
      result.push(text);
    }
  }
  return result;
}

/** Reference metadata on a data-source slot: the c:f formula and, for numRef,
 *  the cached c:formatCode — both round-trip only. */
function readRefMeta(el: XmlElement): { formula?: string; formatCode?: string } {
  const strRef = findChild(el, "c:strRef");
  const numRef = strRef ? undefined : findChild(el, "c:numRef");
  const ref = strRef ?? numRef;
  const meta: { formula?: string; formatCode?: string } = {};
  if (ref) {
    const f = findChild(ref, "c:f");
    const text = f ? textOf(f) : "";
    if (text) meta.formula = text;
  }
  if (numRef) {
    const cache = findChild(numRef, "c:numCache");
    const fc = cache ? findChild(cache, "c:formatCode") : undefined;
    const text = fc ? textOf(fc) : "";
    if (text) meta.formatCode = text;
  }
  // c:numLit carries its formatCode as a direct child
  const numLit = strRef || numRef ? undefined : findChild(el, "c:numLit");
  if (numLit) {
    const fc = findChild(numLit, "c:formatCode");
    const text = fc ? textOf(fc) : "";
    if (text) meta.formatCode = text;
  }
  return meta;
}

function readStrLit(el: XmlElement): string[] | undefined {
  const strLit = findChild(el, "c:strLit");
  if (!strLit) return undefined;
  const result: string[] = [];
  for (const pt of strLit.elements ?? []) {
    if (pt.name === "c:pt") {
      const v = findChild(pt, "c:v");
      const text = v ? textOf(v) : "";
      if (text !== "") result.push(text);
    }
  }
  return result;
}

function readMultiLvlStrCache(el: XmlElement): string[][] | undefined {
  const ref = findChild(el, "c:multiLvlStrRef");
  if (!ref) return undefined;
  const cache = findChild(ref, "c:multiLvlStrCache");
  if (!cache) return [];
  const levels: string[][] = [];
  for (const lvl of cache.elements ?? []) {
    if (lvl.name !== "c:lvl") continue;
    const level: string[] = [];
    for (const pt of lvl.elements ?? []) {
      if (pt.name === "c:pt") {
        const v = findChild(pt, "c:v");
        const text = v ? textOf(v) : "";
        if (text !== "") level.push(text);
      }
    }
    levels.push(level);
  }
  return levels;
}

function readNumLitPoints(numLit: XmlElement): number[] {
  const result: number[] = [];
  for (const pt of numLit.elements ?? []) {
    if (pt.name === "c:pt") {
      const v = findChild(pt, "c:v");
      const text = v ? textOf(v) : "";
      if (text !== "") result.push(Number(text));
    }
  }
  return result;
}

/** A scatter series carries numeric x values when c:xVal holds the
 * CT_NumDataSource choice (numRef/numLit); the label-x shape stringifies
 * x as a strRef category source instead. */
function hasNumericXRef(serEl: XmlElement): boolean {
  const xVal = findChild(serEl, "c:xVal");
  return (
    xVal !== undefined &&
    (findChild(xVal, "c:numRef") !== undefined || findChild(xVal, "c:numLit") !== undefined)
  );
}

function readNumCache(el: XmlElement): number[] {
  // c:numLit literal points (CT_NumDataSource choice: numRef | numLit)
  const numLit = findChild(el, "c:numLit");
  if (numLit) return readNumLitPoints(numLit);
  const numRef = findChild(el, "c:numRef");
  if (!numRef) return [];
  const numCache = findChild(numRef, "c:numCache");
  if (!numCache?.elements) return [];
  const result: number[] = [];
  for (const pt of numCache.elements) {
    if (pt.name === "c:pt") {
      const v = findChild(pt, "c:v");
      const text = v ? textOf(v) : "";
      if (text !== "") result.push(Number(text));
    }
  }
  return result;
}

function readSeriesName(serEl: XmlElement): { name: string | undefined; literal: boolean } {
  const tx = findChild(serEl, "c:tx");
  if (!tx) return { name: undefined, literal: false };
  // CT_SerTx choice: c:strRef (reference) or bare c:v (inline literal).
  const v = findChild(tx, "c:v");
  if (v && !findChild(tx, "c:strRef")) {
    const text = textOf(v);
    return { name: text !== "" ? text : undefined, literal: true };
  }
  const values = readStrCache(tx);
  return { name: values[0] ?? "", literal: false };
}

/** Read a c:title. Plain text stays a string; layout/overlay/spPr/txPr
 *  decorations promote it to the ChartTitleOptions object form. */
/** Read the title's c:tx > c:rich body; undefined when absent or strRef-linked. */
function readTitleBody(
  titleEl: XmlElement,
  ctx: ReadContext,
): string | TextBodyOptions | undefined {
  const rich = findChild(findChild(titleEl, "c:tx") ?? titleEl, "c:rich");
  if (!rich) return undefined;
  const body = textBodyDesc.parse(rich, ctx);
  return collapseTitleBody(body);
}

/**
 * Collapse a parsed title body to its plain text when it is exactly what the
 * string form emits (empty bodyPr/listStyle, one bare run), keeping simple
 * titles readable in the Options JSON.
 */
function collapseTitleBody(body: TextBodyOptions): string | TextBodyOptions {
  if (body.listStyle !== undefined) return body;
  if (body.bodyProperties !== undefined && Object.keys(body.bodyProperties).length > 0) return body;
  const ps = body.paragraphs ?? [];
  if (ps.length !== 1) return body;
  const p = ps[0];
  if (p === undefined) return body;
  if (typeof p === "string") return p;
  if (p.properties !== undefined || p.children !== undefined) return body;
  if (p.endParagraphProperties !== undefined && p.endParagraphProperties !== false) return body;
  return p.text ?? "";
}

function readTitle(titleEl: XmlElement, ctx: ReadContext): string | ChartTitleOptions {
  const body = readTitleBody(titleEl, ctx);
  const text = body === undefined ? readTitleText(titleEl) : undefined;
  const layoutEl = findChild(titleEl, "c:layout");
  const overlayEl = findChild(titleEl, "c:overlay");
  const spPrEl = findChild(titleEl, "c:spPr");
  const txPrEl = findChild(titleEl, "c:txPr");
  // A structured body must wrap as { text } — returning it bare would land a
  // TextBodyOptions where stringify expects ChartTitleOptions and drop c:tx.
  if (!layoutEl && !overlayEl && !spPrEl && !txPrEl) {
    if (typeof body === "string" || body === undefined) return body ?? text ?? "";
    return { text: body };
  }
  const title: ChartTitleOptions = {};
  if (body !== undefined) title.text = body;
  else if (text !== undefined) title.text = text;
  if (layoutEl) {
    const manual = readManualLayout(layoutEl);
    title.layout = manual ?? true;
  }
  if (overlayEl) title.overlay = parseOnOff(attr(overlayEl, "val")) ?? false;
  if (spPrEl)
    title.shapeProperties = shapePropertiesDesc.parse(spPrEl, ctx) as ShapePropertiesOptions;
  if (txPrEl) title.textProperties = textBodyDesc.parse(txPrEl, ctx);
  return title;
}

function readTitleText(titleEl: XmlElement): string | undefined {
  const tx = findChild(titleEl, "c:tx");
  if (!tx?.elements) return undefined;
  for (const child of tx.elements) {
    if (child.name !== "c:rich") continue;
    for (const sub of child.elements ?? []) {
      if (sub.name !== "a:p") continue;
      for (const r of sub.elements ?? []) {
        if (r.name !== "a:r") continue;
        const t = findChild(r, "a:t");
        if (t) {
          const text = textOf(t);
          if (text) return text;
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
  return parseOnOff(v) ?? true;
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
    if (pt.name === "c:pt") {
      const v = findChild(pt, "c:v");
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
    if (type) opts.type = xsdTrendlineType.from(type) as TrendlineOptions["type"];
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
  if (valueType) opts.valueType = xsdErrorValueType.from(valueType) as ErrorBarOptions["valueType"];
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

function readDataLabels(serEl: XmlElement, ctx: ReadContext): DataLabelsOptions | undefined {
  const dlEl = findChild(serEl, "c:dLbls");
  if (!dlEl) return undefined;
  const opts: DataLabelsOptions = {};
  const groupDelete = readBoolAttr(dlEl, "c:delete");
  if (groupDelete !== undefined) opts.delete = groupDelete;
  const numFmt = attr(findChild(dlEl, "c:numFmt"), "formatCode");
  if (numFmt) opts.numberFormat = numFmt;
  const dlSpPr = findChild(dlEl, "c:spPr");
  if (dlSpPr) {
    const shapeProperties = shapePropertiesDesc.parse(dlSpPr, ctx);
    if (shapeProperties) opts.shapeProperties = shapeProperties;
  }
  const dlTxPr = findChild(dlEl, "c:txPr");
  if (dlTxPr) {
    const textProperties = textBodyDesc.parse(dlTxPr, ctx);
    if (textProperties) opts.textProperties = textProperties;
  }
  const position = readValStr(dlEl, "c:dLblPos");
  if (position)
    opts.position = xsdDataLabelPosition.from(position) as DataLabelsOptions["position"];
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
      const labelDelete = readBoolAttr(el, "c:delete");
      if (labelDelete !== undefined) result.delete = labelDelete;
      const layoutEl = findChild(el, "c:layout");
      if (layoutEl) result.layout = readManualLayout(layoutEl) ?? true;
      const richEl = findChild(findChild(el, "c:tx"), "c:rich");
      if (richEl) {
        const text = textBodyDesc.parse(richEl, ctx);
        if (text) result.text = text;
      }
      const numFmt = attr(findChild(el, "c:numFmt"), "formatCode");
      if (numFmt) result.numberFormat = numFmt;
      const lblSpPr = findChild(el, "c:spPr");
      if (lblSpPr) {
        const shapeProperties = shapePropertiesDesc.parse(lblSpPr, ctx);
        if (shapeProperties) result.shapeProperties = shapeProperties;
      }
      const lblTxPr = findChild(el, "c:txPr");
      if (lblTxPr) {
        const textProperties = textBodyDesc.parse(lblTxPr, ctx);
        if (textProperties) result.textProperties = textProperties;
      }
      const pos = readValStr(el, "c:dLblPos");
      if (pos) result.position = xsdDataLabelPosition.from(pos) as DataLabelOptions["position"];
      for (const flag of [
        "showLegendKey",
        "showVal",
        "showCatName",
        "showSerName",
        "showPercent",
        "showBubbleSize",
      ] as const) {
        const v = readBoolAttr(el, `c:${flag}`);
        if (v !== undefined) result[flag] = v;
      }
      const sep = textOf(findChild(el, "c:separator"));
      if (sep) result.separator = sep;
      return result;
    });
  }
  if (findChild(dlEl, "c:leaderLines")) opts.leaderLines = true;
  return opts;
}

function readMarker(parent: XmlElement, ctx: ReadContext): MarkerOptions | undefined {
  const el = findChild(parent, "c:marker");
  if (!el) return undefined;
  const opts: MarkerOptions = {};
  const symbol = readValStr(el, "c:symbol");
  if (symbol) opts.symbol = symbol as MarkerSymbol;
  const size = readValNum(el, "c:size");
  if (size !== undefined) opts.size = size;
  const spPr = findChild(el, "c:spPr");
  if (spPr) {
    const shapeProperties = shapePropertiesDesc.parse(spPr, ctx);
    if (shapeProperties) opts.shapeProperties = shapeProperties;
  }
  return Object.keys(opts).length ? opts : undefined;
}

function readPictureOptions(parent: XmlElement): PictureOptionsOptions | undefined {
  const el = findChild(parent, "c:pictureOptions");
  if (!el) return undefined;
  const opts: PictureOptionsOptions = {};
  const front = readBoolAttr(el, "c:applyToFront");
  if (front !== undefined) opts.applyToFront = front;
  const sides = readBoolAttr(el, "c:applyToSides");
  if (sides !== undefined) opts.applyToSides = sides;
  const end = readBoolAttr(el, "c:applyToEnd");
  if (end !== undefined) opts.applyToEnd = end;
  const fmt = readValStr(el, "c:pictureFormat");
  if (fmt) opts.pictureFormat = fmt as PictureFormat;
  const stack = readValNum(el, "c:pictureStackUnit");
  if (stack !== undefined) opts.pictureStackUnit = stack;
  return Object.keys(opts).length ? opts : undefined;
}

function readDataPoints(serEl: XmlElement, ctx: ReadContext): DataPointOptions[] | undefined {
  const els = children(serEl, "c:dPt");
  if (els.length === 0) return undefined;
  return els.map((el): DataPointOptions => {
    const dp: DataPointOptions = { index: readValNum(el, "c:idx") ?? 0 };
    const invertIfNegative = readBoolAttr(el, "c:invertIfNegative");
    if (invertIfNegative !== undefined) dp.invertIfNegative = invertIfNegative;
    const marker = readMarker(el, ctx);
    if (marker) dp.marker = marker;
    const bubble3D = readBoolAttr(el, "c:bubble3D");
    if (bubble3D !== undefined) dp.bubble3D = bubble3D;
    const explosion = readValNum(el, "c:explosion");
    if (explosion !== undefined) dp.explosion = explosion;
    const spPr = findChild(el, "c:spPr");
    if (spPr) {
      const shapeProperties = shapePropertiesDesc.parse(spPr, ctx);
      if (shapeProperties) dp.shapeProperties = shapeProperties;
    }
    const pictureOptions = readPictureOptions(el);
    if (pictureOptions) dp.pictureOptions = pictureOptions;
    return dp;
  });
}

/** Read shared ser enhancement fields (CT_Ser children beyond EG_SerShared). */
function readSeriesCommon(serEl: XmlElement, ctx: ReadContext): Partial<ChartSeriesCommon> {
  const common: Partial<ChartSeriesCommon> = {};
  // EG_SerShared head: idx/order are unique chart-wide, not array positions —
  // combo charts interleave them across groups, so renumbering by position
  // would produce duplicates Excel treats as corruption.
  const idx = readValNum(serEl, "c:idx");
  if (idx !== undefined) common.index = idx;
  const order = readValNum(serEl, "c:order");
  if (order !== undefined) common.order = order;
  const spPrEl = findChild(serEl, "c:spPr");
  // Presence round-trips: a bare <c:spPr/> placeholder stays an empty object
  // (an XSD-valid form legacy Word writes), not dropped.
  if (spPrEl)
    common.shapeProperties = shapePropertiesDesc.parse(spPrEl, ctx) as ShapePropertiesOptions;
  const trendlines = readTrendlines(serEl);
  if (trendlines) common.trendlines = trendlines;
  const errorBars = readErrBars(serEl);
  if (errorBars) common.errorBars = errorBars;
  const dataLabels = readDataLabels(serEl, ctx);
  if (dataLabels) common.dataLabels = dataLabels;
  const dataPoints = readDataPoints(serEl, ctx);
  if (dataPoints) common.dataPoints = dataPoints;
  const marker = readMarker(serEl, ctx);
  if (marker) common.marker = marker;
  const invertIfNegative = readBoolAttr(serEl, "c:invertIfNegative");
  if (invertIfNegative !== undefined) common.invertIfNegative = invertIfNegative;
  const smooth = readBoolAttr(serEl, "c:smooth");
  if (smooth !== undefined) common.smooth = smooth;
  const explosion = readValNum(serEl, "c:explosion");
  if (explosion !== undefined) common.explosion = explosion;
  const pictureOptions = readPictureOptions(serEl);
  if (pictureOptions) common.pictureOptions = pictureOptions;
  const shape = readValStr(serEl, "c:shape");
  if (shape) common.shape = shape as BarShape;
  const bubble3D = readBoolAttr(serEl, "c:bubble3D");
  if (bubble3D !== undefined) common.bubble3D = bubble3D;
  const extLst = findChild(serEl, "c:extLst");
  if (extLst) {
    const inner = stringifyXml(extLst);
    if (inner) common.ext = inner;
  }
  return common;
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

// ── Plot-area layout read (CT_Layout > CT_ManualLayout) ──

function readManualLayout(layoutEl: XmlElement | undefined): ManualLayoutOptions | undefined {
  if (!layoutEl) return undefined;
  const ml = findChild(layoutEl, "c:manualLayout");
  if (!ml) return undefined;
  const opts: ManualLayoutOptions = {};
  const layoutTarget = readValStr(ml, "c:layoutTarget");
  if (layoutTarget) opts.layoutTarget = layoutTarget as LayoutTarget;
  const xMode = readValStr(ml, "c:xMode");
  if (xMode) opts.xMode = xMode as LayoutMode;
  const yMode = readValStr(ml, "c:yMode");
  if (yMode) opts.yMode = yMode as LayoutMode;
  const wMode = readValStr(ml, "c:wMode");
  if (wMode) opts.wMode = wMode as LayoutMode;
  const hMode = readValStr(ml, "c:hMode");
  if (hMode) opts.hMode = hMode as LayoutMode;
  const x = readValNum(ml, "c:x");
  if (x !== undefined) opts.x = x;
  const y = readValNum(ml, "c:y");
  if (y !== undefined) opts.y = y;
  const w = readValNum(ml, "c:w");
  if (w !== undefined) opts.w = w;
  const h = readValNum(ml, "c:h");
  if (h !== undefined) opts.h = h;
  return Object.keys(opts).length ? opts : undefined;
}

// ── 3D wall/floor surface read (CT_Surface) ──

function readSurface(
  chartEl: XmlElement,
  tag: "c:floor" | "c:sideWall" | "c:backWall",
  ctx: ReadContext,
): SurfaceOptions | undefined {
  const el = findChild(chartEl, tag);
  if (!el) return undefined;
  const opts: SurfaceOptions = {};
  const thickness = readValStr(el, "c:thickness");
  if (thickness !== undefined)
    opts.thickness = thickness.endsWith("%") ? thickness : Number(thickness);
  const spPr = findChild(el, "c:spPr");
  if (spPr) opts.shapeProperties = shapePropertiesDesc.parse(spPr, ctx) as ShapePropertiesOptions;
  return opts;
}

// ── Chart-type scalar read (CT_xxxChart fields between ser and axId) ──

function readChartTypeScalars(
  chartTypeEl: XmlElement | undefined,
  type: ChartType,
  threeD: boolean,
  result: MutableChartSpaceResult,
  ctx: ReadContext,
): void {
  if (!chartTypeEl) return;
  // Group-level varyColors precedes ser in every chart-group content model —
  // except CT_Surface*/CT_Stock, which have no varyColors slot; reading it
  // there would round-trip schema-invalid input into schema-invalid output.
  if (type !== "surface" && type !== "stock") {
    const varyColors = readBoolAttr(chartTypeEl, "c:varyColors");
    if (varyColors !== undefined) result.varyColors = varyColors;
  }
  // c:grouping sits in the header (before ser) of bar/column/line/area groups;
  // only store a non-default value so fresh output keeps the per-type default.
  if (type === "column" || type === "bar" || type === "line" || type === "area") {
    const grouping = readValStr(chartTypeEl, "c:grouping");
    if (grouping) {
      const fallback = type === "column" || type === "bar" ? "clustered" : "standard";
      if (grouping !== fallback) result.grouping = grouping as ChartSpaceOptions["grouping"];
    }
  }
  if (type === "column" || type === "bar") {
    const gapWidth = readValNum(chartTypeEl, "c:gapWidth");
    if (gapWidth !== undefined) result.gapWidth = gapWidth;
    if (threeD) {
      const gapDepth = readValNum(chartTypeEl, "c:gapDepth");
      if (gapDepth !== undefined) result.gapDepth = gapDepth;
      const shape = readValStr(chartTypeEl, "c:shape");
      if (shape) result.shape = shape as BarShape;
    } else {
      const overlap = readValNum(chartTypeEl, "c:overlap");
      if (overlap !== undefined) result.overlap = overlap;
    }
  } else if (type === "area" && threeD) {
    // CT_Area3DChart tail: dropLines → gapDepth
    const gapDepth = readValNum(chartTypeEl, "c:gapDepth");
    if (gapDepth !== undefined) result.gapDepth = gapDepth;
  } else if (type === "pie" || type === "doughnut") {
    const firstSliceAng = readValNum(chartTypeEl, "c:firstSliceAng");
    if (firstSliceAng !== undefined) result.firstSliceAngle = firstSliceAng;
    const holeSize = readValNum(chartTypeEl, "c:holeSize");
    if (holeSize !== undefined) result.holeSize = holeSize;
  } else if (type === "bubble") {
    const bubbleScale = readValNum(chartTypeEl, "c:bubbleScale");
    if (bubbleScale !== undefined) result.bubbleScale = bubbleScale;
    const showNeg = readBoolAttr(chartTypeEl, "c:showNegBubbles");
    if (showNeg !== undefined) result.showNegativeBubbles = showNeg;
    const sizeRep = readValStr(chartTypeEl, "c:sizeRepresents");
    if (sizeRep) result.sizeRepresents = xsdSizeRepresents.from(sizeRep) as SizeRepresents;
  } else if (type === "surface") {
    const wireframe = readBoolAttr(chartTypeEl, "c:wireframe");
    if (wireframe !== undefined) result.wireframe = wireframe;
    const bandFormats = readBandFormats(chartTypeEl);
    if (bandFormats) result.bandFormats = bandFormats;
  } else if (type === "radar") {
    const radarStyle = readValStr(chartTypeEl, "c:radarStyle");
    if (radarStyle) result.radarStyle = radarStyle as RadarStyle;
  } else if (type === "ofPie") {
    const ofPieType = readValStr(chartTypeEl, "c:ofPieType");
    if (ofPieType) result.ofPieType = ofPieType as OfPieType;
    const gapWidth = readValNum(chartTypeEl, "c:gapWidth");
    if (gapWidth !== undefined) result.gapWidth = gapWidth;
    const splitType = readValStr(chartTypeEl, "c:splitType");
    if (splitType) result.splitType = xsdSplitType.from(splitType) as SplitType;
    const splitPos = readValNum(chartTypeEl, "c:splitPos");
    if (splitPos !== undefined) result.splitPosition = splitPos;
    const custSplit = findChild(chartTypeEl, "c:custSplit");
    if (custSplit) {
      const pts = children(custSplit, "c:secondPiePt")
        .map((el) => attr(el, "val"))
        .filter((v): v is string => v !== undefined)
        .map(Number);
      if (pts.length) result.customSplitPoints = pts;
    }
    const secondPieSize = readValStr(chartTypeEl, "c:secondPieSize");
    if (secondPieSize !== undefined)
      result.secondPieSize = secondPieSize.endsWith("%") ? secondPieSize : Number(secondPieSize);
    const seriesLines = readChartLines(chartTypeEl, "c:serLines", ctx);
    if (seriesLines !== undefined) result.seriesLines = seriesLines;
  }
}

// ── Chart-type decoration read (CT_ChartLines containers) ──

function readChartTypeDecorations(
  chartTypeEl: XmlElement | undefined,
  type: ChartType,
  threeD: boolean,
  result: MutableChartSpaceResult,
  ctx: ReadContext,
): void {
  if (!chartTypeEl) return;
  const readUpDownBars = (el: XmlElement) => {
    result.upDownBars = true;
    const gw = readValNum(el, "c:gapWidth");
    if (gw !== undefined) result.upDownBarsGapWidth = gw;
  };
  if (type === "line") {
    if (threeD) {
      const dropLinesV = readChartLines(chartTypeEl, "c:dropLines", ctx);
      if (dropLinesV !== undefined) result.dropLines = dropLinesV;
    } else {
      const highLowLinesV = readChartLines(chartTypeEl, "c:hiLowLines", ctx);
      if (highLowLinesV !== undefined) result.highLowLines = highLowLinesV;
      const upDownBars = findChild(chartTypeEl, "c:upDownBars");
      if (upDownBars) readUpDownBars(upDownBars);
      const dropLinesV = readChartLines(chartTypeEl, "c:dropLines", ctx);
      if (dropLinesV !== undefined) result.dropLines = dropLinesV;
      const markers = readBoolAttr(chartTypeEl, "c:marker");
      if (markers !== undefined) result.markers = markers;
      const smooth = readBoolAttr(chartTypeEl, "c:smooth");
      if (smooth !== undefined) result.smooth = smooth;
    }
  } else if (type === "area") {
    const dropLinesV = readChartLines(chartTypeEl, "c:dropLines", ctx);
    if (dropLinesV !== undefined) result.dropLines = dropLinesV;
    const threeDSeriesLines = readChartLines(chartTypeEl, "c:serLines", ctx);
    if (threeD && threeDSeriesLines !== undefined) result.seriesLines = threeDSeriesLines;
  } else if (type === "stock") {
    const dropLinesV = readChartLines(chartTypeEl, "c:dropLines", ctx);
    if (dropLinesV !== undefined) result.dropLines = dropLinesV;
    const highLowLinesV = readChartLines(chartTypeEl, "c:hiLowLines", ctx);
    if (highLowLinesV !== undefined) result.highLowLines = highLowLinesV;
    const upDownBars = findChild(chartTypeEl, "c:upDownBars");
    if (upDownBars) readUpDownBars(upDownBars);
    const seriesLines = readChartLines(chartTypeEl, "c:serLines", ctx);
    if (seriesLines !== undefined) result.seriesLines = seriesLines;
  } else if ((type === "column" || type === "bar") && !threeD) {
    const seriesLines = readChartLines(chartTypeEl, "c:serLines", ctx);
    if (seriesLines !== undefined) result.seriesLines = seriesLines;
  }
}

// ── Plot-area data table read (CT_DTable) ──

function readDataTable(
  plotArea: XmlElement | undefined,
  ctx: ReadContext,
): DataTableOptions | undefined {
  if (!plotArea) return undefined;
  const dTable = findChild(plotArea, "c:dTable");
  if (!dTable) return undefined;
  const opts: DataTableOptions = {};
  const horz = readBoolAttr(dTable, "c:showHorzBorder");
  if (horz !== undefined) opts.showHorizontalBorder = horz;
  const vert = readBoolAttr(dTable, "c:showVertBorder");
  if (vert !== undefined) opts.showVerticalBorder = vert;
  const outline = readBoolAttr(dTable, "c:showOutline");
  if (outline !== undefined) opts.showOutline = outline;
  const keys = readBoolAttr(dTable, "c:showKeys");
  if (keys !== undefined) opts.showLegendKeys = keys;
  const spPrEl = findChild(dTable, "c:spPr");
  if (spPrEl)
    opts.shapeProperties = shapePropertiesDesc.parse(spPrEl, ctx) as ShapePropertiesOptions;
  const txPrEl = findChild(dTable, "c:txPr");
  if (txPrEl) opts.textProperties = textBodyDesc.parse(txPrEl, ctx);
  // Presence round-trips: a bare <c:dTable/> stays an empty object.
  return opts;
}

// ── ChartSpace-level read (CT_ChartSpace children) ──

function readProtection(el: XmlElement): ProtectionOptions | undefined {
  const protection = findChild(el, "c:protection");
  if (!protection) return undefined;
  const opts: ProtectionOptions = {};
  const readFlag = (tag: string): boolean | undefined => {
    const child = findChild(protection, tag);
    if (!child) return undefined;
    const v = attr(child, "val");
    return v === undefined ? undefined : (parseOnOff(v) ?? true);
  };
  const chartObject = readFlag("c:chartObject");
  if (chartObject !== undefined) opts.chartObject = chartObject;
  const data = readFlag("c:data");
  if (data !== undefined) opts.data = data;
  const formatting = readFlag("c:formatting");
  if (formatting !== undefined) opts.formatting = formatting;
  const selection = readFlag("c:selection");
  if (selection !== undefined) opts.selection = selection;
  const userInterface = readFlag("c:userInterface");
  if (userInterface !== undefined) opts.userInterface = userInterface;
  return Object.keys(opts).length ? opts : undefined;
}

function readExternalData(el: XmlElement): ExternalDataOptions | undefined {
  const externalData = findChild(el, "c:externalData");
  if (!externalData) return undefined;
  const relationshipId = attr(externalData, "r:id");
  if (relationshipId === undefined) return undefined;
  const opts: ExternalDataOptions = { relationshipId };
  const autoUpdateEl = findChild(externalData, "c:autoUpdate");
  if (autoUpdateEl) {
    const v = attr(autoUpdateEl, "val");
    if (v !== undefined) opts.autoUpdate = parseOnOff(v) ?? true;
  }
  return opts;
}

function readPivotSource(el: XmlElement): PivotSourceOptions | undefined {
  const ps = findChild(el, "c:pivotSource");
  if (!ps) return undefined;
  const nameEl = findChild(ps, "c:name");
  const fmtIdEl = findChild(ps, "c:fmtId");
  if (!nameEl || !fmtIdEl) return undefined;
  return { name: textOf(nameEl) ?? "", formatId: Number(attr(fmtIdEl, "val")) };
}

function readPivotFormats(
  chart: XmlElement,
  ctx: ReadContext,
): ChartPivotFormatOptions[] | undefined {
  const pf = findChild(chart, "c:pivotFmts");
  if (!pf) return undefined;
  const result: ChartPivotFormatOptions[] = [];
  for (const fmt of children(pf, "c:pivotFmt")) {
    const idxEl = findChild(fmt, "c:idx");
    if (!idxEl) continue;
    const opt: ChartPivotFormatOptions = { index: Number(attr(idxEl, "val")) };
    const marker = readMarker(fmt, ctx);
    if (marker) opt.marker = marker;
    result.push(opt);
  }
  return result.length ? result : undefined;
}

function readBandFormats(chartTypeEl: XmlElement): BandFormatOptions[] | undefined {
  const bf = findChild(chartTypeEl, "c:bandFmts");
  if (!bf) return undefined;
  const result: BandFormatOptions[] = [];
  for (const fmt of children(bf, "c:bandFmt")) {
    const idxEl = findChild(fmt, "c:idx");
    if (idxEl) result.push({ index: Number(attr(idxEl, "val")) });
  }
  return result.length ? result : undefined;
}

function readLegendEntries(legend: XmlElement, ctx: ReadContext): LegendEntryOptions[] | undefined {
  const entries = children(legend, "c:legendEntry");
  if (!entries.length) return undefined;
  const result: LegendEntryOptions[] = [];
  for (const entry of entries) {
    const idxEl = findChild(entry, "c:idx");
    if (!idxEl) continue;
    const deleteEl = findChild(entry, "c:delete");
    if (deleteEl) {
      const v = attr(deleteEl, "val");
      result.push({
        index: Number(attr(idxEl, "val")),
        delete: parseOnOff(v) ?? true,
      });
      continue;
    }
    const txPrEl = findChild(entry, "c:txPr");
    if (txPrEl) {
      result.push({
        index: Number(attr(idxEl, "val")),
        textProperties: textBodyDesc.parse(txPrEl, ctx) as TextBodyOptions,
      });
    }
  }
  return result.length ? result : undefined;
}

function readLegend(
  chart: XmlElement,
  ctx: ReadContext,
):
  | {
      position?: LegendPosition;
      entries?: LegendEntryOptions[];
      layout?: boolean | ManualLayoutOptions;
      overlay?: boolean;
      shapeProperties?: ShapePropertiesOptions;
      textProperties?: TextBodyOptions;
    }
  | undefined {
  const legend = findChild(chart, "c:legend");
  if (!legend) return undefined;
  const result: {
    position?: LegendPosition;
    entries?: LegendEntryOptions[];
    layout?: boolean | ManualLayoutOptions;
    overlay?: boolean;
    shapeProperties?: ShapePropertiesOptions;
    textProperties?: TextBodyOptions;
  } = {};
  const posEl = findChild(legend, "c:legendPos");
  if (posEl) {
    const v = attr(posEl, "val");
    if (v) result.position = xsdLegendPosition.from(v) as LegendPosition;
  }
  const entries = readLegendEntries(legend, ctx);
  if (entries) result.entries = entries;
  const layoutEl = findChild(legend, "c:layout");
  if (layoutEl) result.layout = readManualLayout(layoutEl) ?? true;
  const overlayEl = findChild(legend, "c:overlay");
  if (overlayEl) result.overlay = parseOnOff(attr(overlayEl, "val")) ?? false;
  const spPrEl = findChild(legend, "c:spPr");
  if (spPrEl)
    result.shapeProperties = shapePropertiesDesc.parse(spPrEl, ctx) as ShapePropertiesOptions;
  const txPr = findChild(legend, "c:txPr");
  if (txPr) result.textProperties = textBodyDesc.parse(txPr, ctx);
  return Object.keys(result).length ? result : undefined;
}

// ── Print settings read (CT_PrintSettings) ──

function readHeaderFooter(ps: XmlElement): ChartHeaderFooterOptions | undefined {
  const hf = findChild(ps, "c:headerFooter");
  if (!hf) return undefined;
  const opts: ChartHeaderFooterOptions = {};
  const readStr = (tag: string): string | undefined => {
    const child = findChild(hf, tag);
    return child ? textOf(child) : undefined;
  };
  const oh = readStr("c:oddHeader");
  if (oh !== undefined) opts.oddHeader = oh;
  const of = readStr("c:oddFooter");
  if (of !== undefined) opts.oddFooter = of;
  const eh = readStr("c:evenHeader");
  if (eh !== undefined) opts.evenHeader = eh;
  const ef = readStr("c:evenFooter");
  if (ef !== undefined) opts.evenFooter = ef;
  const fh = readStr("c:firstHeader");
  if (fh !== undefined) opts.firstHeader = fh;
  const ff = readStr("c:firstFooter");
  if (ff !== undefined) opts.firstFooter = ff;
  const readFlag = (name: string): boolean | undefined => {
    const v = attr(hf, name);
    return v === undefined ? undefined : (parseOnOff(v) ?? true);
  };
  const alignWithMargins = readFlag("alignWithMargins");
  if (alignWithMargins !== undefined) opts.alignWithMargins = alignWithMargins;
  const differentOddEven = readFlag("differentOddEven");
  if (differentOddEven !== undefined) opts.differentOddEven = differentOddEven;
  const differentFirst = readFlag("differentFirst");
  if (differentFirst !== undefined) opts.differentFirst = differentFirst;
  // An empty <c:headerFooter/> round-trips as {} — presence is the payload.
  return opts;
}

function readPageMargins(ps: XmlElement): ChartPageMarginsOptions | undefined {
  const pm = findChild(ps, "c:pageMargins");
  if (!pm) return undefined;
  const opts: ChartPageMarginsOptions = {};
  const readNum = (name: string): number | undefined => {
    const v = attr(pm, name);
    return v === undefined ? undefined : Number(v);
  };
  const left = readNum("l");
  if (left !== undefined) opts.left = left;
  const right = readNum("r");
  if (right !== undefined) opts.right = right;
  const top = readNum("t");
  if (top !== undefined) opts.top = top;
  const bottom = readNum("b");
  if (bottom !== undefined) opts.bottom = bottom;
  const header = readNum("header");
  if (header !== undefined) opts.header = header;
  const footer = readNum("footer");
  if (footer !== undefined) opts.footer = footer;
  return Object.keys(opts).length ? opts : undefined;
}

function readPageSetup(ps: XmlElement): ChartPageSetupOptions | undefined {
  const pg = findChild(ps, "c:pageSetup");
  if (!pg) return undefined;
  const opts: ChartPageSetupOptions = {};
  const readNum = (name: string): number | undefined => {
    const v = attr(pg, name);
    return v === undefined ? undefined : Number(v);
  };
  const readFlag = (name: string): boolean | undefined => {
    const v = attr(pg, name);
    return v === undefined ? undefined : (parseOnOff(v) ?? true);
  };
  const paperSize = readNum("paperSize");
  if (paperSize !== undefined) opts.paperSize = paperSize;
  const paperHeight = attr(pg, "paperHeight");
  if (paperHeight !== undefined) opts.paperHeight = paperHeight;
  const paperWidth = attr(pg, "paperWidth");
  if (paperWidth !== undefined) opts.paperWidth = paperWidth;
  const firstPageNumber = readNum("firstPageNumber");
  if (firstPageNumber !== undefined) opts.firstPageNumber = firstPageNumber;
  const orientation = attr(pg, "orientation");
  if (orientation !== undefined) opts.orientation = orientation as PageSetupOrientation;
  const blackAndWhite = readFlag("blackAndWhite");
  if (blackAndWhite !== undefined) opts.blackAndWhite = blackAndWhite;
  const draft = readFlag("draft");
  if (draft !== undefined) opts.draft = draft;
  const useFirstPageNumber = readFlag("useFirstPageNumber");
  if (useFirstPageNumber !== undefined) opts.useFirstPageNumber = useFirstPageNumber;
  const horizontalDpi = readNum("horizontalDpi");
  if (horizontalDpi !== undefined) opts.horizontalDpi = horizontalDpi;
  const verticalDpi = readNum("verticalDpi");
  if (verticalDpi !== undefined) opts.verticalDpi = verticalDpi;
  const copies = readNum("copies");
  if (copies !== undefined) opts.copies = copies;
  // An empty <c:pageSetup/> round-trips as {} — presence is the payload.
  return opts;
}

function readPrintSettings(el: XmlElement): PrintSettingsOptions | undefined {
  const ps = findChild(el, "c:printSettings");
  if (!ps) return undefined;
  const opts: PrintSettingsOptions = {};
  const headerFooter = readHeaderFooter(ps);
  if (headerFooter) opts.headerFooter = headerFooter;
  const pageMargins = readPageMargins(ps);
  if (pageMargins) opts.pageMargins = pageMargins;
  const pageSetup = readPageSetup(ps);
  if (pageSetup) opts.pageSetup = pageSetup;
  const legacyDrawingHF = findChild(ps, "c:legacyDrawingHF");
  if (legacyDrawingHF) {
    const rId = attr(legacyDrawingHF, "r:id");
    if (rId !== undefined) opts.legacyDrawingId = rId;
  }
  return Object.keys(opts).length ? opts : undefined;
}

// ── Axis read (CT_CatAx / CT_ValAx / CT_DateAx / CT_SerAx) ──

const AXIS_KIND_BY_TAG: Record<string, AxisKind> = {
  "c:catAx": "category",
  "c:valAx": "value",
  "c:dateAx": "date",
  "c:serAx": "series",
};

function readAxis(el: XmlElement, kind: AxisKind, ctx: ReadContext): AxisOptions {
  const result: AxisOptions = {
    kind,
    id: readValNum(el, "c:axId") ?? 0,
    crossAxisId: readValNum(el, "c:crossAx") ?? 0,
  };

  const scalingEl = findChild(el, "c:scaling");
  if (scalingEl) {
    const scaling: AxisScalingOptions = {};
    const logBase = readValNum(scalingEl, "c:logBase");
    if (logBase !== undefined) scaling.logBase = logBase;
    const orientation = readValStr(scalingEl, "c:orientation");
    if (orientation) scaling.orientation = xsdAxisOrientation.from(orientation) as AxisOrientation;
    const max = readValNum(scalingEl, "c:max");
    if (max !== undefined) scaling.max = max;
    const min = readValNum(scalingEl, "c:min");
    if (min !== undefined) scaling.min = min;
    if (logBase !== undefined || orientation || max !== undefined || min !== undefined) {
      result.scaling = scaling;
    }
  }

  const del = readBoolAttr(el, "c:delete");
  if (del !== undefined) result.delete = del;
  const pos = readValStr(el, "c:axPos");
  if (pos) result.position = xsdAxisPosition.from(pos) as AxisPosition;
  result.majorGridlines = readChartLines(el, "c:majorGridlines", ctx);
  result.minorGridlines = readChartLines(el, "c:minorGridlines", ctx);
  const titleEl = findChild(el, "c:title");
  if (titleEl) {
    // "" keeps a text-less title placeholder (bare <c:title/>) round-tripping.
    result.title = readTitle(titleEl, ctx);
  }
  const numFmtEl = findChild(el, "c:numFmt");
  if (numFmtEl) {
    const formatCode = attr(numFmtEl, "formatCode");
    if (formatCode) {
      const sourceLinked = attr(numFmtEl, "sourceLinked");
      // Only a decoupled sourceLinked (0) needs the object form; the default
      // 1 round-trips through the plain-string shorthand.
      result.numberFormat = sourceLinked === "0" ? { formatCode, sourceLinked: false } : formatCode;
    }
  }
  const majorTickMark = readValStr(el, "c:majorTickMark");
  if (majorTickMark) result.majorTickMark = majorTickMark as AxisTickMark;
  const minorTickMark = readValStr(el, "c:minorTickMark");
  if (minorTickMark) result.minorTickMark = minorTickMark as AxisTickMark;
  const tickLblPos = readValStr(el, "c:tickLblPos");
  if (tickLblPos) result.tickLabelPosition = tickLblPos as AxisTickLabelPosition;
  // EG_AxShared: spPr sits between tickLblPos and crossAx; txPr follows spPr.
  const spPrEl = findChild(el, "c:spPr");
  if (spPrEl) {
    result.shapeProperties = shapePropertiesDesc.parse(spPrEl, ctx) as ShapePropertiesOptions;
  }
  const axisTxPr = findChild(el, "c:txPr");
  if (axisTxPr) result.textProperties = textBodyDesc.parse(axisTxPr, ctx);
  const crossesAt = readValNum(el, "c:crossesAt");
  if (crossesAt !== undefined) result.crossesAt = crossesAt;
  else {
    const crosses = readValStr(el, "c:crosses");
    if (crosses) result.crosses = xsdAxisCrosses.from(crosses) as AxisCrosses;
  }

  // kind-specific tail (mirrors stringifyAxis switch order)
  if (kind === "category" || kind === "date") {
    const auto = readBoolAttr(el, "c:auto");
    if (auto !== undefined) result.auto = auto;
  }
  if (kind === "category") {
    const lblAlgn = readValStr(el, "c:lblAlgn");
    if (lblAlgn) result.labelAlignment = xsdAxisLabelAlignment.from(lblAlgn) as AxisLabelAlignment;
  }
  if (kind === "category" || kind === "date") {
    const lblOffsetEl = findChild(el, "c:lblOffset");
    if (lblOffsetEl) {
      const v = attr(lblOffsetEl, "val");
      if (v !== undefined) result.labelOffset = v.endsWith("%") ? v : Number(v);
    }
  }
  if (kind === "category" || kind === "series") {
    const tickLblSkip = readValNum(el, "c:tickLblSkip");
    if (tickLblSkip !== undefined) result.tickLabelSkip = tickLblSkip;
    const tickMarkSkip = readValNum(el, "c:tickMarkSkip");
    if (tickMarkSkip !== undefined) result.tickMarkSkip = tickMarkSkip;
  }
  if (kind === "category") {
    const noMulti = readBoolAttr(el, "c:noMultiLvlLbl");
    if (noMulti !== undefined) result.noMultiLevelLabel = noMulti;
  }
  if (kind === "date") {
    const baseTimeUnit = readValStr(el, "c:baseTimeUnit");
    if (baseTimeUnit) result.baseTimeUnit = baseTimeUnit as TimeUnit;
    const majorTimeUnit = readValStr(el, "c:majorTimeUnit");
    if (majorTimeUnit) result.majorTimeUnit = majorTimeUnit as TimeUnit;
    const minorTimeUnit = readValStr(el, "c:minorTimeUnit");
    if (minorTimeUnit) result.minorTimeUnit = minorTimeUnit as TimeUnit;
  }
  if (kind === "date" || kind === "value") {
    const majorUnit = readValNum(el, "c:majorUnit");
    if (majorUnit !== undefined) result.majorUnit = majorUnit;
    const minorUnit = readValNum(el, "c:minorUnit");
    if (minorUnit !== undefined) result.minorUnit = minorUnit;
  }
  if (kind === "value") {
    const crossBetween = readValStr(el, "c:crossBetween");
    if (crossBetween)
      result.crossBetween = xsdAxisCrossBetween.from(crossBetween) as AxisCrossBetween;
    const dispUnitsEl = findChild(el, "c:dispUnits");
    if (dispUnitsEl) {
      const displayUnits: DisplayUnitsOptions = {};
      const customUnit = readValNum(dispUnitsEl, "c:custUnit");
      if (customUnit !== undefined) displayUnits.customUnit = customUnit;
      const builtInUnit = readValStr(dispUnitsEl, "c:builtInUnit");
      if (builtInUnit) displayUnits.builtInUnit = builtInUnit as BuiltInDisplayUnit;
      if (findChild(dispUnitsEl, "c:dispUnitsLbl")) displayUnits.label = true;
      if (customUnit !== undefined || builtInUnit || displayUnits.label) {
        result.displayUnits = displayUnits;
      }
    }
  }
  return result;
}

/** Read all axis elements (c:catAx/c:valAx/c:dateAx/c:serAx) under plotArea. */
function readAxes(plotArea: XmlElement | undefined, ctx: ReadContext): AxisOptions[] | undefined {
  if (!plotArea?.elements) return undefined;
  const axes: AxisOptions[] = [];
  for (const child of plotArea.elements) {
    if (!child.name) continue;
    const kind = AXIS_KIND_BY_TAG[child.name];
    if (kind) axes.push(readAxis(child, kind, ctx));
  }
  return axes.length > 0 ? axes : undefined;
}

// ── Main descriptor ──

export const chartSpaceDesc: CustomDescriptor<ChartSpaceOptions> = {
  kind: "custom",

  stringify(opts: ChartSpaceOptions, ctx: WriteContext): string {
    const parts: string[] = [];

    // Opening tag with namespaces — the strict dialect re-declares the
    // purl.oclc.org set a strict source package requires.
    const ns =
      opts.dialect === "strict"
        ? `xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart" xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"`
        : `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
    parts.push(`<c:chartSpace ${ns}>`);

    // Optional header elements — emitted only when set (round-trip from a
    // source that carried them). Fresh output omits them, matching the
    // legacy-Word corpus form.
    if (opts.date1904 !== undefined) parts.push(`<c:date1904${boolVal(opts.date1904)}/>`);
    if (opts.lang !== undefined) parts.push(valEl("c:lang", opts.lang));
    if (opts.roundedCorners !== undefined)
      parts.push(`<c:roundedCorners${boolVal(opts.roundedCorners)}/>`);

    // Style — plain c:style, or the Word 2010+ mc:AlternateContent form
    // whose mc:Fallback carries the equivalent c:style.
    if (opts.style2010) {
      const fallback =
        opts.style2010.fallbackStyle !== undefined
          ? `<mc:Fallback>${valEl("c:style", opts.style2010.fallbackStyle)}</mc:Fallback>`
          : "";
      parts.push(
        `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
          `<mc:Choice xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart" Requires="c14">` +
          `<c14:style val="${opts.style2010.style}"/></mc:Choice>` +
          `${fallback}</mc:AlternateContent>`,
      );
    } else if (opts.style !== undefined) {
      parts.push(valEl("c:style", opts.style));
    }

    // CT_ChartSpace: style → clrMapOvr → pivotSource → protection → chart
    if (opts.colorMappingOverride) {
      parts.push(stringifyColorMapping(opts.colorMappingOverride, "c:clrMapOvr"));
    }
    if (opts.pivotSource) parts.push(stringifyPivotSource(opts.pivotSource));
    if (opts.protection) parts.push(stringifyProtection(opts.protection));

    // c:chart container
    parts.push("<c:chart>");

    // Title — a non-empty string emits the c:rich form; "" emits the bare
    // auto-title placeholder legacy Word writes.
    if (opts.title !== undefined) {
      parts.push(
        typeof opts.title === "string" && opts.title === ""
          ? emptyEl("c:title")
          : stringifyTitle(opts.title, ctx),
      );
    }
    if (opts.autoTitleDeleted !== undefined)
      parts.push(`<c:autoTitleDeleted${boolVal(opts.autoTitleDeleted)}/>`);
    // CT_Chart: autoTitleDeleted → pivotFmts → view3D
    if (opts.pivotFormats?.length) parts.push(stringifyPivotFormats(opts.pivotFormats, ctx));

    // 3D view (before plotArea per CT_Chart sequence)
    if (opts.view3D) {
      parts.push(stringifyView3D(opts.view3D));
    }

    // 3D walls and floor (CT_Chart: view3D → floor → sideWall → backWall → plotArea)
    parts.push(stringifySurface("c:floor", opts.floor, ctx));
    parts.push(stringifySurface("c:sideWall", opts.sideWall, ctx));
    parts.push(stringifySurface("c:backWall", opts.backWall, ctx));

    // c:plotArea
    parts.push(`<c:plotArea>${stringifyLayout(opts.plotAreaLayout)}`);

    // Chart type element (header + series + footer)
    parts.push(chartTypeHeader(opts));

    // A parsed empty chart shell carries no series (parse omits the field
    // when the source had no c:ser) — emit none instead of crashing.
    for (const [i, series] of (opts.series ?? []).entries()) {
      parts.push(stringifySeries(i, series, opts, ctx));
    }

    // Chart-group-level data labels follow the ser elements (EG_*ChartShared).
    if (opts.dataLabels) parts.push(stringifyDataLabels(opts.dataLabels, ctx));

    parts.push(chartTypeFooter(opts, ctx));

    // Combo chart: a second group (e.g. lines over bars) shares the plot
    // area, the category source, and the axes list.
    for (const g of opts.secondaryGroups ?? []) parts.push(stringifySecondaryGroup(g, opts, ctx));

    // Axes
    parts.push(stringifyAxes(opts, ctx));

    // Data table (CT_PlotArea: axes → dTable → spPr)
    if (opts.dataTable) parts.push(stringifyDataTable(opts.dataTable, ctx));
    parts.push(chartSpPr(opts.plotAreaShapeProperties, ctx));

    parts.push("</c:plotArea>");

    // Legend
    if (opts.showLegend !== false) {
      parts.push(stringifyLegend(opts, ctx));
    }

    // CT_Chart tail: plotVisOnly → dispBlanksAs → showDLblsOverMax
    parts.push(`<c:plotVisOnly${boolVal(opts.plotVisOnly ?? true)}/>`);
    if (opts.displayBlanksAs !== undefined)
      parts.push(valEl("c:dispBlanksAs", opts.displayBlanksAs));
    if (opts.showDataLabelsOverMax !== undefined)
      parts.push(`<c:showDLblsOverMax${boolVal(opts.showDataLabelsOverMax)}/>`);

    parts.push("</c:chart>");

    // Chart-area spPr — optional, emitted only when the source carried one.
    parts.push(chartSpPr(opts.shapeProperties, ctx));

    // CT_ChartSpace: txPr → externalData → printSettings → (userShapes)
    if (opts.textProperties) {
      parts.push(`<c:txPr>${textBodyDesc.stringify(opts.textProperties, ctx) ?? ""}</c:txPr>`);
    }
    if (opts.externalData) {
      const ext = stringifyExternalData(opts.externalData);
      if (ext) parts.push(ext);
    }
    if (opts.printSettings) parts.push(stringifyPrintSettings(opts.printSettings));
    // CT_ChartSpace: printSettings → userShapes → extLst
    if (opts.userShapes !== undefined)
      parts.push(`<c:userShapes r:id="${escapeXml(opts.userShapes.relationshipId ?? "rId1")}"/>`);
    if (opts.ext) parts.push(`<c:extLst>${opts.ext}</c:extLst>`);

    parts.push("</c:chartSpace>");

    return parts.join("");
  },

  parse(el: XmlElement, ctx: ReadContext) {
    const result: MutableChartSpaceResult = {};

    // Namespace dialect — a strict root declares the purl.oclc.org chart
    // namespace; everything else reads as transitional.
    if (el.attributes?.["xmlns:c"] === "http://purl.oclc.org/ooxml/drawingml/chart") {
      result.dialect = "strict";
    }

    // Optional header elements — recorded only when present so stringify
    // re-emits exactly what the source carried.
    const date1904 = readBoolAttr(el, "c:date1904");
    if (date1904 !== undefined) result.date1904 = date1904;
    const langVal = readValStr(el, "c:lang");
    if (langVal) result.lang = langVal;
    const roundedCorners = readBoolAttr(el, "c:roundedCorners");
    if (roundedCorners !== undefined) result.roundedCorners = roundedCorners;

    // Style — bare c:style, or the Word 2010+ mc:AlternateContent form whose
    // mc:Choice holds c14:style and mc:Fallback the equivalent c:style.
    const mcStyleEl = findChild(el, "mc:AlternateContent");
    const c14StyleEl =
      mcStyleEl && findChild(findChild(mcStyleEl, "mc:Choice") ?? mcStyleEl, "c14:style");
    if (c14StyleEl?.attributes?.["val"] !== undefined) {
      const fallbackEl = findFirst(mcStyleEl, "c:style");
      result.style2010 = {
        style: Number(c14StyleEl.attributes["val"]),
        ...(fallbackEl?.attributes?.["val"] !== undefined
          ? { fallbackStyle: Number(fallbackEl.attributes["val"]) }
          : {}),
      };
    } else {
      const styleEl = findChild(el, "c:style");
      if (styleEl?.attributes?.["val"] !== undefined) {
        result.style = Number(styleEl.attributes["val"]);
      }
    }

    // CT_ChartSpace: clrMapOvr + protection (before c:chart)
    const colorMappingOverride = parseColorMapping(findChild(el, "c:clrMapOvr"));
    if (colorMappingOverride) result.colorMappingOverride = colorMappingOverride;
    const protection = readProtection(el);
    if (protection) result.protection = protection;
    const pivotSource = readPivotSource(el);
    if (pivotSource) result.pivotSource = pivotSource;

    // Chart container
    const chart = findChild(el, "c:chart");
    if (!chart) return result as ChartSpaceOptions;

    // Title — "" records a text-less title placeholder (bare <c:title/>) so
    // presence round-trips; autoTitleDeleted follows the same rule.
    const titleEl = findChild(chart, "c:title");
    if (titleEl) {
      result.title = readTitle(titleEl, ctx);
    }
    const autoTitleDeleted = readBoolAttr(chart, "c:autoTitleDeleted");
    if (autoTitleDeleted !== undefined) result.autoTitleDeleted = autoTitleDeleted;

    // 3D view (before plotArea per CT_Chart sequence)
    const pivotFormats = readPivotFormats(chart, ctx);
    if (pivotFormats) result.pivotFormats = pivotFormats;
    const view3D = readView3D(chart);
    if (view3D) result.view3D = view3D;

    // 3D walls and floor
    const floor = readSurface(chart, "c:floor", ctx);
    if (floor) result.floor = floor;
    const sideWall = readSurface(chart, "c:sideWall", ctx);
    if (sideWall) result.sideWall = sideWall;
    const backWall = readSurface(chart, "c:backWall", ctx);
    if (backWall) result.backWall = backWall;

    // Plot area — detect chart type
    const plotArea = findChild(chart, "c:plotArea");
    if (plotArea?.elements) {
      // plot-area manual layout (CT_Layout > CT_ManualLayout)
      const manualLayout = readManualLayout(findChild(plotArea, "c:layout"));
      if (manualLayout) result.plotAreaLayout = manualLayout;
      // plot-area data table (CT_PlotArea tail)
      const dataTable = readDataTable(plotArea, ctx);
      if (dataTable) result.dataTable = dataTable;
      let detectedType: ChartType | undefined;
      let threeD = false;
      let chartTypeEl: XmlElement | undefined;
      // Combo charts carry further chart groups in the same plot area
      // (CT_PlotArea allows any number of them).
      const secondaryEls: XmlElement[] = [];

      for (const child of plotArea.elements) {
        if (!child.name) continue;
        const mapping = TAG_TO_CHART_TYPE[child.name];
        if (mapping) {
          if (chartTypeEl) {
            secondaryEls.push(child);
            continue;
          }
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
        }
      }

      if (detectedType && chartTypeEl) {
        result.type = detectedType;
        if (threeD) result.threeD = true;
        readChartTypeScalars(chartTypeEl, detectedType, threeD, result, ctx);
        readChartTypeDecorations(chartTypeEl, detectedType, threeD, result, ctx);
        // Chart-group-level data labels (after the ser elements).
        const groupLabels = readDataLabels(chartTypeEl, ctx);
        if (groupLabels) result.dataLabels = groupLabels;
        // The axId reference sequence is data in its own right (legacy files
        // carry dangling val="0" entries) — record it verbatim.
        const axIdEls = children(chartTypeEl, "c:axId");
        if (axIdEls.length > 0) {
          result.axisIds = axIdEls.map((e) => Number(attr(e, "val"))).filter((n) => !isNaN(n));
        }
      }

      // Extract series data (c:ser are children of the chart-type element, not plotArea)
      const seriesEls = chartTypeEl ? children(chartTypeEl, "c:ser") : [];
      if (seriesEls.length > 0) {
        if (detectedType === "bubble") {
          const bubbleSeries: BubbleSeriesData[] = [];
          for (const serEl of seriesEls) {
            const { name } = readSeriesName(serEl);
            const xVal = findChild(serEl, "c:xVal");
            const yVal = findChild(serEl, "c:yVal");
            const bubbleSize = findChild(serEl, "c:bubbleSize");
            bubbleSeries.push({
              name,
              xValues: xVal ? readNumCache(xVal) : [],
              yValues: yVal ? readNumCache(yVal) : [],
              bubbleSize: bubbleSize ? readNumCache(bubbleSize) : [],
              ...readSeriesCommon(serEl, ctx),
            });
          }
          result.series = bubbleSeries as BubbleSeriesData[];
        } else if (detectedType === "scatter" && hasNumericXRef(seriesEls[0]!)) {
          // Numeric-axis scatter (c:xVal/c:yVal as CT_NumDataSource) — per-series
          // xy arrays; the label-x shape falls through to the category path.
          const xySeries: ScatterSeriesData[] = [];
          for (const serEl of seriesEls) {
            const { name, literal: nameLiteral } = readSeriesName(serEl);
            const txEl = findChild(serEl, "c:tx");
            const nameFormula = txEl ? readRefMeta(txEl).formula : undefined;
            const xVal = findChild(serEl, "c:xVal");
            const yVal = findChild(serEl, "c:yVal");
            const valueLiteral = yVal ? findChild(yVal, "c:numLit") !== undefined : false;
            const yMeta = yVal ? readRefMeta(yVal) : {};
            xySeries.push({
              name,
              ...(nameLiteral ? { nameLiteral } : {}),
              ...(nameFormula ? { nameFormula } : {}),
              ...(valueLiteral ? { valueLiteral } : {}),
              ...(yMeta.formula ? { valueFormula: yMeta.formula } : {}),
              ...(yMeta.formatCode ? { formatCode: yMeta.formatCode } : {}),
              xValues: xVal ? readNumCache(xVal) : [],
              yValues: yVal ? readNumCache(yVal) : [],
              ...readSeriesCommon(serEl, ctx),
            });
          }
          result.series = xySeries;
        } else {
          const chartSeries: ChartSeriesData[] = [];
          let categories: string[] | undefined;

          for (const serEl of seriesEls) {
            const { name, literal: nameLiteral } = readSeriesName(serEl);
            const txEl = findChild(serEl, "c:tx");
            const nameFormula = txEl ? readRefMeta(txEl).formula : undefined;

            // Read categories from first series (c:cat may be strRef/strLit/
            // multiLvlStrRef/numRef — numeric categories keep their literal text)
            if (!categories && !result.multiLevelCategories && !result.categoryLabels) {
              const catEl = findChild(serEl, "c:cat") ?? findChild(serEl, "c:xVal");
              if (catEl) {
                const multi = readMultiLvlStrCache(catEl);
                if (multi) {
                  result.multiLevelCategories = multi;
                  const multiRef = findChild(catEl, "c:multiLvlStrRef");
                  const multiF = multiRef ? findChild(multiRef, "c:f") : undefined;
                  const multiText = multiF ? textOf(multiF) : "";
                  if (multiText) result.categoryFormula = multiText;
                } else {
                  const lit = readStrLit(catEl);
                  if (lit) {
                    result.categoryLabels = lit;
                    categories = lit;
                  } else if (findChild(catEl, "c:numRef")) {
                    // A numRef source is numeric even when the cache is empty
                    // (ptCount=0) — the axis kind is the reference form, not
                    // the cached points.
                    result.numericCategories = true;
                    categories = readNumCacheText(catEl);
                  } else {
                    categories = readStrCache(catEl);
                  }
                  const catMeta = readRefMeta(catEl);
                  if (catMeta.formula) result.categoryFormula = catMeta.formula;
                  if (catMeta.formatCode) result.categoryFormatCode = catMeta.formatCode;
                }
              }
            }

            // Read values
            const valueEl = findChild(serEl, "c:val") ?? findChild(serEl, "c:yVal");
            const valueLiteral = valueEl ? findChild(valueEl, "c:numLit") !== undefined : false;
            const values = valueEl ? readNumCache(valueEl) : [];
            const valMeta = valueEl ? readRefMeta(valueEl) : {};

            chartSeries.push({
              name,
              ...(nameLiteral ? { nameLiteral } : {}),
              ...(nameFormula ? { nameFormula } : {}),
              ...(valueLiteral ? { valueLiteral } : {}),
              ...(valMeta.formula ? { valueFormula: valMeta.formula } : {}),
              ...(valMeta.formatCode ? { formatCode: valMeta.formatCode } : {}),
              values,
              ...readSeriesCommon(serEl, ctx),
            });
          }

          result.series = chartSeries as ChartSeriesData[];
          if (categories) result.categories = categories;
        }
      }

      // Further chart groups (combo charts) — own series and group flags,
      // sharing the category source and the axes list with the main group.
      if (secondaryEls.length > 0) {
        const groups: SecondaryChartGroupOptions[] = [];
        for (const secondaryEl of secondaryEls) {
          const g = readSecondaryGroup(secondaryEl, ctx);
          if (g) groups.push(g);
        }
        if (groups.length > 0) result.secondaryGroups = groups;
      }
    }

    // Axes
    const axes = readAxes(plotArea, ctx);
    if (axes) result.axes = axes;

    // CT_PlotArea tail: dTable → spPr
    const plotSpPr = plotArea && findChild(plotArea, "c:spPr");
    if (plotSpPr)
      result.plotAreaShapeProperties = shapePropertiesDesc.parse(
        plotSpPr,
        ctx,
      ) as ShapePropertiesOptions;

    // Legend — default is true; parse records true explicitly so stringify
    // can tell a round-trip legend from a fresh default one.
    const legend = findChild(chart, "c:legend");
    if (!legend) {
      result.showLegend = false;
    } else {
      result.showLegend = true;
      const legendData = readLegend(chart, ctx);
      if (legendData?.position) result.legendPosition = legendData.position;
      if (legendData?.layout !== undefined) result.legendLayout = legendData.layout;
      if (legendData?.overlay !== undefined) result.legendOverlay = legendData.overlay;
      if (legendData?.shapeProperties) result.legendShapeProperties = legendData.shapeProperties;
      if (legendData?.textProperties) result.legendTextProperties = legendData.textProperties;
      if (legendData?.entries) result.legendEntries = legendData.entries;
    }

    // CT_Chart tail: plotVisOnly, dispBlanksAs, showDLblsOverMax
    const plotVisOnly = readBoolAttr(chart, "c:plotVisOnly");
    if (plotVisOnly !== undefined) result.plotVisOnly = plotVisOnly;
    const dispBlanksAs = readValStr(chart, "c:dispBlanksAs");
    if (dispBlanksAs) result.displayBlanksAs = dispBlanksAs as DisplayBlanksAs;
    const showDLblsOverMax = readBoolAttr(chart, "c:showDLblsOverMax");
    if (showDLblsOverMax !== undefined) result.showDataLabelsOverMax = showDLblsOverMax;

    // CT_ChartSpace tail: spPr (chart-area shape properties), txPr, then externalData
    const spPrEl = findChild(el, "c:spPr");
    if (spPrEl) {
      result.shapeProperties = shapePropertiesDesc.parse(spPrEl, ctx) as ShapePropertiesOptions;
    }
    const spaceTxPr = findChild(el, "c:txPr");
    if (spaceTxPr) result.textProperties = textBodyDesc.parse(spaceTxPr, ctx);
    const externalData = readExternalData(el);
    if (externalData) result.externalData = externalData;
    const printSettings = readPrintSettings(el);
    if (printSettings) result.printSettings = printSettings;
    const userShapesEl = findChild(el, "c:userShapes");
    if (userShapesEl) {
      const rid = attr(userShapesEl, "r:id");
      // anchors stay empty here — the format parsers resolve the companion
      // part body from the chart part's own rels and fill it in
      if (rid !== undefined) result.userShapes = { relationshipId: rid, anchors: [] };
    }
    const spaceExtLst = findChild(el, "c:extLst");
    if (spaceExtLst) {
      result.ext = (spaceExtLst.elements ?? []).map((e) => stringifyElement(e)).join("");
    }

    return result as ChartSpaceOptions;
  },
};
