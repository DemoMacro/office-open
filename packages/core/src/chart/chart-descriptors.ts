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
  ManualLayoutOptions,
  SurfaceOptions,
  LayoutTarget,
  LayoutMode,
  DisplayBlanksAs,
  SizeRepresents,
  DataTableOptions,
  OfPieType,
  SplitType,
  ColorMapOverrideOptions,
  ColorSchemeIndex,
  ProtectionOptions,
  ExternalDataOptions,
  HeaderFooterOptions,
  PageMarginsOptions,
  PageSetupOptions,
  PrintSettingsOptions,
  PageSetupOrientation,
  PivotSourceOptions,
  PivotFormatOptions,
  BandFormatOptions,
  LegendEntryOptions,
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
): string {
  if (!opts) return "";
  const parts: string[] = [];
  if (opts.thickness !== undefined) parts.push(valEl("c:thickness", opts.thickness));
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
  if (opts.orientation !== undefined) parts.push(valEl("c:orientation", opts.orientation));
  if (opts.max !== undefined) parts.push(valEl("c:max", opts.max));
  if (opts.min !== undefined) parts.push(valEl("c:min", opts.min));
  return `<c:scaling>${parts.join("")}</c:scaling>`;
}

function stringifyAxisTitle(title: string): string {
  // Same c:rich run shape as the chart title so parse reuses readTitleText.
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeXml(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
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
function stringifyAxis(opts: AxisOptions): string {
  const parts: string[] = [];
  parts.push(valEl("c:axId", opts.id));
  parts.push(stringifyScaling(opts.scaling));
  if (opts.delete !== undefined) parts.push(`<c:delete${boolVal(opts.delete)}/>`);
  if (opts.position !== undefined) parts.push(valEl("c:axPos", opts.position));
  if (opts.majorGridlines) parts.push(emptyEl("c:majorGridlines"));
  if (opts.minorGridlines) parts.push(emptyEl("c:minorGridlines"));
  if (opts.title !== undefined) parts.push(stringifyAxisTitle(opts.title));
  if (opts.numberFormat !== undefined)
    parts.push(`<c:numFmt formatCode="${escapeXml(opts.numberFormat)}" sourceLinked="1"/>`);
  if (opts.majorTickMark !== undefined) parts.push(valEl("c:majorTickMark", opts.majorTickMark));
  if (opts.minorTickMark !== undefined) parts.push(valEl("c:minorTickMark", opts.minorTickMark));
  if (opts.tickLabelPosition !== undefined)
    parts.push(valEl("c:tickLblPos", opts.tickLabelPosition));
  parts.push(valEl("c:crossAx", opts.crossAxisId));
  // XSD choice: crosses XOR crossesAt
  if (opts.crossesAt !== undefined) parts.push(valEl("c:crossesAt", opts.crossesAt));
  else if (opts.crosses !== undefined) parts.push(valEl("c:crosses", opts.crosses));

  switch (opts.kind) {
    case "category":
      if (opts.auto !== undefined) parts.push(`<c:auto${boolVal(opts.auto)}/>`);
      if (opts.labelAlignment !== undefined) parts.push(valEl("c:lblAlgn", opts.labelAlignment));
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
      if (opts.crossBetween !== undefined) parts.push(valEl("c:crossBetween", opts.crossBetween));
      if (opts.majorUnit !== undefined) parts.push(valEl("c:majorUnit", opts.majorUnit));
      if (opts.minorUnit !== undefined) parts.push(valEl("c:minorUnit", opts.minorUnit));
      if (opts.displayUnits) parts.push(stringifyDisplayUnits(opts.displayUnits));
      break;
  }
  const tag = AXIS_TAG[opts.kind];
  return `<${tag}>${parts.join("")}</${tag}>`;
}

// ── Series data XML builders ──

function stringifyStrRef(values: readonly string[]): string {
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${escapeXml(v)}</c:v></c:pt>`).join("");
  return `<c:strRef><c:f/><c:strCache><c:ptCount ${attrVal("val", values.length)}/>${pts}</c:strCache></c:strRef>`;
}

function stringifyStrLit(values: readonly string[]): string {
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${escapeXml(v)}</c:v></c:pt>`).join("");
  return `<c:strLit><c:ptCount ${attrVal("val", values.length)}/>${pts}</c:strLit>`;
}

function stringifyMultiLvlStrRef(levels: readonly (readonly string[])[]): string {
  const ptCount = levels.reduce((max, lvl) => Math.max(max, lvl.length), 0);
  const lvls = levels
    .map((lvl) => {
      const pts = lvl.map((v, i) => `<c:pt idx="${i}"><c:v>${escapeXml(v)}</c:v></c:pt>`).join("");
      return `<c:lvl>${pts}</c:lvl>`;
    })
    .join("");
  return `<c:multiLvlStrRef><c:f/><c:multiLvlStrCache><c:ptCount ${attrVal("val", ptCount)}/>${lvls}</c:multiLvlStrCache></c:multiLvlStrRef>`;
}

// CT_AxDataSource choice (c:cat / c:xVal slot): exactly one of multi-level/literal/reference.
function stringifyCategorySource(opts: ChartSpaceOptions): string {
  if (opts.multiLevelCategories) return stringifyMultiLvlStrRef(opts.multiLevelCategories);
  if (opts.categoryLabels) return stringifyStrLit(opts.categoryLabels);
  return stringifyStrRef(opts.categories ?? []);
}

function stringifyNumRef(values: readonly number[]): string {
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("");
  return `<c:numRef><c:f/><c:numCache><c:formatCode>General</c:formatCode><c:ptCount ${attrVal("val", values.length)}/>${pts}</c:numCache></c:numRef>`;
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

function stringifyDataTable(opts: DataTableOptions): string {
  // CT_DTable: showHorzBorder → showVertBorder → showOutline → showKeys
  const parts: string[] = [];
  if (opts.showHorizontalBorder !== undefined)
    parts.push(`<c:showHorzBorder${boolVal(opts.showHorizontalBorder)}/>`);
  if (opts.showVerticalBorder !== undefined)
    parts.push(`<c:showVertBorder${boolVal(opts.showVerticalBorder)}/>`);
  if (opts.showOutline !== undefined) parts.push(`<c:showOutline${boolVal(opts.showOutline)}/>`);
  if (opts.showLegendKeys !== undefined) parts.push(`<c:showKeys${boolVal(opts.showLegendKeys)}/>`);
  return `<c:dTable>${parts.join("")}</c:dTable>`;
}

const COLOR_MAP_KEYS = [
  "bg1",
  "tx1",
  "bg2",
  "tx2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
] as const;

const COLOR_MAP_IDENTITY: Required<ColorMapOverrideOptions> = {
  bg1: "lt1",
  tx1: "dk1",
  bg2: "lt2",
  tx2: "dk2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hlink",
  folHlink: "folHlink",
};

function stringifyColorMapOverride(opts: ColorMapOverrideOptions): string {
  // CT_ColorMapping: 12 ST_ColorSchemeIndex attributes, all XSD use="required".
  // Merge caller overrides onto the identity mapping so emission is always complete.
  const merged = { ...COLOR_MAP_IDENTITY, ...opts };
  const attrs = COLOR_MAP_KEYS.map((k) => `${k}="${escapeXml(merged[k])}"`).join(" ");
  return `<c:clrMapOvr ${attrs}/>`;
}

function stringifyProtection(opts: ProtectionOptions): string {
  const parts: string[] = [];
  if (opts.chartObject !== undefined)
    parts.push(`<c:chartObject val="${opts.chartObject ? 1 : 0}"/>`);
  if (opts.data !== undefined) parts.push(`<c:data val="${opts.data ? 1 : 0}"/>`);
  if (opts.formatting !== undefined) parts.push(`<c:formatting val="${opts.formatting ? 1 : 0}"/>`);
  if (opts.selection !== undefined) parts.push(`<c:selection val="${opts.selection ? 1 : 0}"/>`);
  if (opts.userInterface !== undefined)
    parts.push(`<c:userInterface val="${opts.userInterface ? 1 : 0}"/>`);
  return `<c:protection>${parts.join("")}</c:protection>`;
}

function stringifyExternalData(opts: ExternalDataOptions): string {
  const auto =
    opts.autoUpdate !== undefined ? `<c:autoUpdate val="${opts.autoUpdate ? 1 : 0}"/>` : "";
  return `<c:externalData r:id="${escapeXml(opts.relationshipId)}">${auto}</c:externalData>`;
}

function stringifyHeaderFooter(opts: HeaderFooterOptions): string {
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
  return `<c:headerFooter${attrStr}>${parts.join("")}</c:headerFooter>`;
}

function stringifyPageMargins(opts: PageMarginsOptions): string {
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

function stringifyPageSetup(opts: PageSetupOptions): string {
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
    case "ofPie":
      headerParts.push(valEl("c:ofPieType", opts.ofPieType ?? "pie"));
      headerParts.push(valEl("c:varyColors", 1));
      break;
    case "pie":
    case "doughnut":
    case "bubble":
      headerParts.push(valEl("c:varyColors", 1));
      break;
    case "surface":
      // EG_SurfaceChartShared head: wireframe precedes ser.
      if (opts.wireframe !== undefined)
        headerParts.push(`<c:wireframe${boolVal(opts.wireframe)}/>`);
      break;
  }

  return `<${tag}>${headerParts.join("")}`;
}

function chartTypeFooter(opts: ChartSpaceOptions): string {
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
      } else {
        // CT_BarChart: gapWidth → overlap → serLines
        if (opts.overlap !== undefined) parts.push(valEl("c:overlap", opts.overlap));
        if (opts.seriesLines) parts.push(emptyEl("c:serLines"));
      }
      break;
    case "line":
      // EG_LineChartShared tail: dropLines (after ser/dLbls)
      if (opts.dropLines) parts.push(emptyEl("c:dropLines"));
      if (!opts.threeD) {
        // CT_LineChart: dropLines → hiLowLines → upDownBars → marker → smooth
        if (opts.highLowLines) parts.push(emptyEl("c:hiLowLines"));
        if (opts.upDownBars) parts.push(stringifyUpDownBars(opts.upDownBarsGapWidth));
      } else if (opts.gapDepth !== undefined) {
        // CT_Line3DChart: dropLines → gapDepth
        parts.push(valEl("c:gapDepth", opts.gapDepth));
      }
      break;
    case "area":
      // EG_AreaChartShared tail: dropLines
      if (opts.dropLines) parts.push(emptyEl("c:dropLines"));
      if (opts.threeD && opts.gapDepth !== undefined)
        parts.push(valEl("c:gapDepth", opts.gapDepth));
      break;
    case "stock":
      // CT_StockChart: dropLines → hiLowLines → upDownBars
      if (opts.dropLines) parts.push(emptyEl("c:dropLines"));
      if (opts.highLowLines) parts.push(emptyEl("c:hiLowLines"));
      if (opts.upDownBars) parts.push(stringifyUpDownBars(opts.upDownBarsGapWidth));
      break;
    case "pie":
      // CT_PieChart: firstSliceAng (no holeSize)
      if (opts.firstSliceAngle !== undefined)
        parts.push(valEl("c:firstSliceAng", opts.firstSliceAngle));
      break;
    case "doughnut":
      // CT_DoughnutChart: firstSliceAng → holeSize
      if (opts.firstSliceAngle !== undefined)
        parts.push(valEl("c:firstSliceAng", opts.firstSliceAngle));
      if (opts.holeSize !== undefined) parts.push(valEl("c:holeSize", opts.holeSize));
      break;
    case "bubble":
      // CT_BubbleChart: bubbleScale → showNegBubbles → sizeRepresents
      if (opts.bubbleScale !== undefined) parts.push(valEl("c:bubbleScale", opts.bubbleScale));
      if (opts.showNegativeBubbles !== undefined)
        parts.push(`<c:showNegBubbles${boolVal(opts.showNegativeBubbles)}/>`);
      if (opts.sizeRepresents !== undefined)
        parts.push(valEl("c:sizeRepresents", opts.sizeRepresents));
      break;
    case "ofPie":
      // CT_OfPieChart: gapWidth → splitType → splitPos → custSplit → secondPieSize → serLines
      if (opts.gapWidth !== undefined) parts.push(valEl("c:gapWidth", opts.gapWidth));
      if (opts.splitType !== undefined) parts.push(valEl("c:splitType", opts.splitType));
      if (opts.splitPosition !== undefined) parts.push(valEl("c:splitPos", opts.splitPosition));
      if (opts.customSplitPoints?.length) {
        const pts = opts.customSplitPoints.map((p) => valEl("c:secondPiePt", p)).join("");
        parts.push(`<c:custSplit>${pts}</c:custSplit>`);
      }
      if (opts.secondPieSize !== undefined)
        parts.push(valEl("c:secondPieSize", opts.secondPieSize));
      if (opts.seriesLines) parts.push(emptyEl("c:serLines"));
      break;
    case "surface":
      // EG_SurfaceChartShared tail: bandFmts (after ser, before axId)
      if (opts.bandFormats?.length) parts.push(stringifyBandFormats(opts.bandFormats));
      break;
  }

  // Axes (pie/doughnut have none)
  if (!NO_AXES_TYPES.has(opts.type)) {
    parts.push(valEl("c:axId", 10));
    parts.push(valEl("c:axId", 20));
    if (opts.threeD || opts.type === "surface") parts.push(valEl("c:axId", 30));
  }

  return `${parts.join("")}</${tag}>`;
}

// ── Series sub-elements (CT_Marker / CT_DPt / CT_PictureOptions) ──

function stringifyMarker(opts: MarkerOptions): string {
  const parts: string[] = [];
  if (opts.symbol !== undefined) parts.push(valEl("c:symbol", opts.symbol));
  if (opts.size !== undefined) parts.push(valEl("c:size", opts.size));
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

function stringifyDataPoint(opts: DataPointOptions): string {
  // CT_DPt: idx → invertIfNegative → marker → bubble3D → explosion → spPr → pictureOptions
  const parts: string[] = [valEl("c:idx", opts.index)];
  if (opts.invertIfNegative !== undefined)
    parts.push(`<c:invertIfNegative${boolVal(opts.invertIfNegative)}/>`);
  if (opts.marker) parts.push(stringifyMarker(opts.marker));
  if (opts.bubble3D !== undefined) parts.push(`<c:bubble3D${boolVal(opts.bubble3D)}/>`);
  if (opts.explosion !== undefined) parts.push(valEl("c:explosion", opts.explosion));
  if (opts.pictureOptions) parts.push(stringifyPictureOptions(opts.pictureOptions));
  return `<c:dPt>${parts.join("")}</c:dPt>`;
}

// ── Series XML ──

function stringifySeries(
  index: number,
  series: ChartSeriesData | BubbleSeriesData,
  opts: ChartSpaceOptions,
): string {
  const chartType = opts.type;
  const categories = opts.categories ?? [];
  const parts: string[] = [];
  const s = series as ChartSeriesData;

  // EG_SerShared: idx, order, tx, spPr
  parts.push(valEl("c:idx", index));
  parts.push(valEl("c:order", index));
  parts.push(`<c:tx>${stringifyStrRef([series.name])}</c:tx>`);
  parts.push(emptyEl("c:spPr"));

  // type-specific head before dPt (per CT_xxxSer content model)
  if (chartType === "line" || chartType === "scatter" || chartType === "radar") {
    if (s.marker) parts.push(stringifyMarker(s.marker));
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
    for (const dp of s.dataPoints) parts.push(stringifyDataPoint(dp));
  }
  if (s.dataLabels) parts.push(stringifyDataLabels(s.dataLabels));
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
    parts.push(`<c:xVal>${stringifyStrRef(categories)}</c:xVal>`);
    parts.push(`<c:yVal>${stringifyNumRef(s.values)}</c:yVal>`);
  } else {
    parts.push(`<c:cat>${stringifyCategorySource(opts)}</c:cat>`);
    parts.push(`<c:val>${stringifyNumRef(s.values)}</c:val>`);
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

  return `<c:ser>${parts.join("")}</c:ser>`;
}

// ── Title XML ──

function stringifyTitle(title: string): string {
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeXml(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
}

// ── Legend XML ──

const LEGEND_TXPR =
  '<c:txPr><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>';

function stringifyLegendEntry(entry: LegendEntryOptions): string {
  // CT_LegendEntry: idx → choice(delete | EG_LegendEntryData[txPr]). Non-delete
  // entries require a txPr model (chart txPr wiring is a follow-up), so only
  // delete entries are emitted here.
  if (!entry.delete) return "";
  return `<c:legendEntry><c:idx val="${entry.index}"/><c:delete val="1"/></c:legendEntry>`;
}

function stringifyLegend(opts: ChartSpaceOptions): string {
  const pos = opts.legendPosition ?? "b";
  const entries = (opts.legendEntries ?? []).map(stringifyLegendEntry).join("");
  return `<c:legend><c:legendPos val="${pos}"/>${entries}<c:layout/><c:overlay val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>${LEGEND_TXPR}</c:legend>`;
}

function stringifyPivotSource(opts: PivotSourceOptions): string {
  return `<c:pivotSource><c:name>${escapeXml(opts.name)}</c:name><c:fmtId val="${opts.formatId}"/></c:pivotSource>`;
}

function stringifyPivotFormats(formats: readonly PivotFormatOptions[]): string {
  const inner = formats
    .map((f) => {
      let entry = `<c:idx val="${f.index}"/>`;
      if (f.marker) entry += stringifyMarker(f.marker);
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

function noFillSpPr(): string {
  return `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>`;
}

function chartTxPr(): string {
  return `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
}

// ── Axes XML ──

/** Default bottom category axis (c:catAx) for generated charts. */
function categoryAxis(id: number, crossAxisId: number): AxisOptions {
  return {
    kind: "category",
    id,
    crossAxisId,
    scaling: { orientation: "minMax" },
    delete: false,
    position: "b",
    crosses: "autoZero",
    auto: true,
    labelOffset: 100,
    noMultiLevelLabel: false,
  };
}

/** Default left value axis (c:valAx). */
function valueAxis(id: number, crossAxisId: number): AxisOptions {
  return {
    kind: "value",
    id,
    crossAxisId,
    scaling: { orientation: "minMax" },
    delete: false,
    position: "l",
    numberFormat: "General",
    crosses: "autoZero",
  };
}

/** Default series axis (c:serAx) for surface charts. */
function seriesAxis(id: number, crossAxisId: number): AxisOptions {
  return {
    kind: "series",
    id,
    crossAxisId,
    scaling: { orientation: "minMax" },
    delete: false,
    position: "b",
    numberFormat: "General",
    crosses: "autoZero",
    tickLabelSkip: 1,
  };
}

/** Sensible default axes derived from chart type, matching prior hardcoded output. */
function defaultAxesFor(chartType: ChartType, threeD?: boolean): readonly AxisOptions[] {
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
    return [categoryAxis(10, 20), valueAxis(20, 10), categoryAxis(30, 10)];
  }
  return [categoryAxis(10, 20), valueAxis(20, 10)];
}

/** Provided axes override defaults; otherwise defaults are derived from chart type. */
function stringifyAxes(opts: ChartSpaceOptions): string {
  const axes = opts.axes ?? defaultAxesFor(opts.type, opts.threeD);
  return axes.map(stringifyAxis).join("");
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
      const text = v ? textOf(v) : "";
      if (text !== "") result.push(text);
    }
  }
  return result;
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

function readNumCache(el: XmlElement): number[] {
  const numRef = findChild(el, "c:numRef");
  if (!numRef) return [];
  const numCache = findChild(numRef, "c:numCache");
  if (!numCache?.elements) return [];
  const result: number[] = [];
  for (const pt of numCache.elements) {
    if (pt.name === "c:pt" && pt.elements) {
      const v = pt.elements.find((c) => c.name === "c:v");
      const text = v ? textOf(v) : "";
      if (text !== "") result.push(Number(text));
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
                if (t.name === "a:t") {
                  const text = textOf(t);
                  if (text) return text;
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

function readMarker(parent: XmlElement): MarkerOptions | undefined {
  const el = findChild(parent, "c:marker");
  if (!el) return undefined;
  const opts: MarkerOptions = {};
  const symbol = readValStr(el, "c:symbol");
  if (symbol) opts.symbol = symbol as MarkerSymbol;
  const size = readValNum(el, "c:size");
  if (size !== undefined) opts.size = size;
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

function readDataPoints(serEl: XmlElement): DataPointOptions[] | undefined {
  const els = children(serEl, "c:dPt");
  if (els.length === 0) return undefined;
  return els.map((el): DataPointOptions => {
    const dp: DataPointOptions = { index: readValNum(el, "c:idx") ?? 0 };
    const invertIfNegative = readBoolAttr(el, "c:invertIfNegative");
    if (invertIfNegative !== undefined) dp.invertIfNegative = invertIfNegative;
    const marker = readMarker(el);
    if (marker) dp.marker = marker;
    const bubble3D = readBoolAttr(el, "c:bubble3D");
    if (bubble3D !== undefined) dp.bubble3D = bubble3D;
    const explosion = readValNum(el, "c:explosion");
    if (explosion !== undefined) dp.explosion = explosion;
    const pictureOptions = readPictureOptions(el);
    if (pictureOptions) dp.pictureOptions = pictureOptions;
    return dp;
  });
}

/** Read shared ser enhancement fields (CT_Ser children beyond EG_SerShared). */
function readSeriesCommon(serEl: XmlElement): Partial<ChartSeriesCommon> {
  const common: Partial<ChartSeriesCommon> = {};
  const trendlines = readTrendlines(serEl);
  if (trendlines) common.trendlines = trendlines;
  const errorBars = readErrBars(serEl);
  if (errorBars) common.errorBars = errorBars;
  const dataLabels = readDataLabels(serEl);
  if (dataLabels) common.dataLabels = dataLabels;
  const dataPoints = readDataPoints(serEl);
  if (dataPoints) common.dataPoints = dataPoints;
  const marker = readMarker(serEl);
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
): SurfaceOptions | undefined {
  const el = findChild(chartEl, tag);
  if (!el) return undefined;
  const thickness = readValStr(el, "c:thickness");
  if (thickness === undefined) return undefined;
  return { thickness: thickness.endsWith("%") ? thickness : Number(thickness) };
}

// ── Chart-type scalar read (CT_xxxChart fields between ser and axId) ──

function readChartTypeScalars(
  chartTypeEl: XmlElement | undefined,
  type: ChartType,
  threeD: boolean,
  result: MutableChartSpaceResult,
): void {
  if (!chartTypeEl) return;
  if (type === "column" || type === "bar") {
    const gapWidth = readValNum(chartTypeEl, "c:gapWidth");
    if (gapWidth !== undefined) result.gapWidth = gapWidth;
    if (threeD) {
      const gapDepth = readValNum(chartTypeEl, "c:gapDepth");
      if (gapDepth !== undefined) result.gapDepth = gapDepth;
    } else {
      const overlap = readValNum(chartTypeEl, "c:overlap");
      if (overlap !== undefined) result.overlap = overlap;
    }
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
    if (sizeRep) result.sizeRepresents = sizeRep as SizeRepresents;
  } else if (type === "surface") {
    const wireframe = readBoolAttr(chartTypeEl, "c:wireframe");
    if (wireframe !== undefined) result.wireframe = wireframe;
    const bandFormats = readBandFormats(chartTypeEl);
    if (bandFormats) result.bandFormats = bandFormats;
  } else if (type === "ofPie") {
    const ofPieType = readValStr(chartTypeEl, "c:ofPieType");
    if (ofPieType) result.ofPieType = ofPieType as OfPieType;
    const gapWidth = readValNum(chartTypeEl, "c:gapWidth");
    if (gapWidth !== undefined) result.gapWidth = gapWidth;
    const splitType = readValStr(chartTypeEl, "c:splitType");
    if (splitType) result.splitType = splitType as SplitType;
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
    if (findChild(chartTypeEl, "c:serLines")) result.seriesLines = true;
  }
}

// ── Chart-type decoration read (CT_ChartLines containers) ──

function readChartTypeDecorations(
  chartTypeEl: XmlElement | undefined,
  type: ChartType,
  threeD: boolean,
  result: MutableChartSpaceResult,
): void {
  if (!chartTypeEl) return;
  const readUpDownBars = (el: XmlElement) => {
    result.upDownBars = true;
    const gw = readValNum(el, "c:gapWidth");
    if (gw !== undefined) result.upDownBarsGapWidth = gw;
  };
  if (type === "line") {
    if (threeD) {
      if (findChild(chartTypeEl, "c:dropLines")) result.dropLines = true;
    } else {
      if (findChild(chartTypeEl, "c:hiLowLines")) result.highLowLines = true;
      const upDownBars = findChild(chartTypeEl, "c:upDownBars");
      if (upDownBars) readUpDownBars(upDownBars);
      if (findChild(chartTypeEl, "c:dropLines")) result.dropLines = true;
    }
  } else if (type === "area") {
    if (findChild(chartTypeEl, "c:dropLines")) result.dropLines = true;
    if (threeD && findChild(chartTypeEl, "c:serLines")) result.seriesLines = true;
  } else if (type === "stock") {
    if (findChild(chartTypeEl, "c:dropLines")) result.dropLines = true;
    if (findChild(chartTypeEl, "c:hiLowLines")) result.highLowLines = true;
    const upDownBars = findChild(chartTypeEl, "c:upDownBars");
    if (upDownBars) readUpDownBars(upDownBars);
    if (findChild(chartTypeEl, "c:serLines")) result.seriesLines = true;
  } else if ((type === "column" || type === "bar") && !threeD) {
    if (findChild(chartTypeEl, "c:serLines")) result.seriesLines = true;
  }
}

// ── Plot-area data table read (CT_DTable) ──

function readDataTable(plotArea: XmlElement | undefined): DataTableOptions | undefined {
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
  return Object.keys(opts).length ? opts : undefined;
}

// ── ChartSpace-level read (CT_ChartSpace children) ──

function readColorMapOverride(el: XmlElement): ColorMapOverrideOptions | undefined {
  const clrMapOvr = findChild(el, "c:clrMapOvr");
  if (!clrMapOvr?.attributes) return undefined;
  const opts: ColorMapOverrideOptions = {};
  for (const k of COLOR_MAP_KEYS) {
    const v = clrMapOvr.attributes[k];
    if (v !== undefined) opts[k] = v as ColorSchemeIndex;
  }
  return Object.keys(opts).length ? opts : undefined;
}

function readProtection(el: XmlElement): ProtectionOptions | undefined {
  const protection = findChild(el, "c:protection");
  if (!protection) return undefined;
  const opts: ProtectionOptions = {};
  const readFlag = (tag: string): boolean | undefined => {
    const child = findChild(protection, tag);
    if (!child) return undefined;
    const v = attr(child, "val");
    return v === undefined ? undefined : v !== "0" && v !== "false";
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
    if (v !== undefined) opts.autoUpdate = v !== "0" && v !== "false";
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

function readPivotFormats(chart: XmlElement): PivotFormatOptions[] | undefined {
  const pf = findChild(chart, "c:pivotFmts");
  if (!pf) return undefined;
  const result: PivotFormatOptions[] = [];
  for (const fmt of children(pf, "c:pivotFmt")) {
    const idxEl = findChild(fmt, "c:idx");
    if (!idxEl) continue;
    const opt: PivotFormatOptions = { index: Number(attr(idxEl, "val")) };
    const marker = readMarker(fmt);
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

function readLegendEntries(legend: XmlElement): LegendEntryOptions[] | undefined {
  const entries = children(legend, "c:legendEntry");
  if (!entries.length) return undefined;
  const result: LegendEntryOptions[] = [];
  for (const entry of entries) {
    const idxEl = findChild(entry, "c:idx");
    if (!idxEl) continue;
    const deleteEl = findChild(entry, "c:delete");
    if (!deleteEl) continue; // non-delete entries (txPr) need chart txPr wiring
    const v = attr(deleteEl, "val");
    result.push({
      index: Number(attr(idxEl, "val")),
      delete: v === undefined ? true : v !== "0",
    });
  }
  return result.length ? result : undefined;
}

function readLegend(
  chart: XmlElement,
): { position?: LegendPosition; entries?: LegendEntryOptions[] } | undefined {
  const legend = findChild(chart, "c:legend");
  if (!legend) return undefined;
  const result: { position?: LegendPosition; entries?: LegendEntryOptions[] } = {};
  const posEl = findChild(legend, "c:legendPos");
  if (posEl) {
    const v = attr(posEl, "val");
    if (v) result.position = v as LegendPosition;
  }
  const entries = readLegendEntries(legend);
  if (entries) result.entries = entries;
  return Object.keys(result).length ? result : undefined;
}

// ── Print settings read (CT_PrintSettings) ──

function readHeaderFooter(ps: XmlElement): HeaderFooterOptions | undefined {
  const hf = findChild(ps, "c:headerFooter");
  if (!hf) return undefined;
  const opts: HeaderFooterOptions = {};
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
    return v === undefined ? undefined : v !== "0" && v !== "false";
  };
  const alignWithMargins = readFlag("alignWithMargins");
  if (alignWithMargins !== undefined) opts.alignWithMargins = alignWithMargins;
  const differentOddEven = readFlag("differentOddEven");
  if (differentOddEven !== undefined) opts.differentOddEven = differentOddEven;
  const differentFirst = readFlag("differentFirst");
  if (differentFirst !== undefined) opts.differentFirst = differentFirst;
  return Object.keys(opts).length ? opts : undefined;
}

function readPageMargins(ps: XmlElement): PageMarginsOptions | undefined {
  const pm = findChild(ps, "c:pageMargins");
  if (!pm) return undefined;
  const opts: PageMarginsOptions = {};
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

function readPageSetup(ps: XmlElement): PageSetupOptions | undefined {
  const pg = findChild(ps, "c:pageSetup");
  if (!pg) return undefined;
  const opts: PageSetupOptions = {};
  const readNum = (name: string): number | undefined => {
    const v = attr(pg, name);
    return v === undefined ? undefined : Number(v);
  };
  const readFlag = (name: string): boolean | undefined => {
    const v = attr(pg, name);
    return v === undefined ? undefined : v !== "0" && v !== "false";
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
  return Object.keys(opts).length ? opts : undefined;
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

function readAxis(el: XmlElement, kind: AxisKind): AxisOptions {
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
    if (orientation) scaling.orientation = orientation as AxisOrientation;
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
  if (pos) result.position = pos as AxisPosition;
  if (findChild(el, "c:majorGridlines")) result.majorGridlines = true;
  if (findChild(el, "c:minorGridlines")) result.minorGridlines = true;
  const titleEl = findChild(el, "c:title");
  if (titleEl) {
    const titleText = readTitleText(titleEl);
    if (titleText !== undefined) result.title = titleText;
  }
  const numFmtEl = findChild(el, "c:numFmt");
  if (numFmtEl) {
    const formatCode = attr(numFmtEl, "formatCode");
    if (formatCode) result.numberFormat = formatCode;
  }
  const majorTickMark = readValStr(el, "c:majorTickMark");
  if (majorTickMark) result.majorTickMark = majorTickMark as AxisTickMark;
  const minorTickMark = readValStr(el, "c:minorTickMark");
  if (minorTickMark) result.minorTickMark = minorTickMark as AxisTickMark;
  const tickLblPos = readValStr(el, "c:tickLblPos");
  if (tickLblPos) result.tickLabelPosition = tickLblPos as AxisTickLabelPosition;
  const crossesAt = readValNum(el, "c:crossesAt");
  if (crossesAt !== undefined) result.crossesAt = crossesAt;
  else {
    const crosses = readValStr(el, "c:crosses");
    if (crosses) result.crosses = crosses as AxisCrosses;
  }

  // kind-specific tail (mirrors stringifyAxis switch order)
  if (kind === "category" || kind === "date") {
    const auto = readBoolAttr(el, "c:auto");
    if (auto !== undefined) result.auto = auto;
  }
  if (kind === "category") {
    const lblAlgn = readValStr(el, "c:lblAlgn");
    if (lblAlgn) result.labelAlignment = lblAlgn as AxisLabelAlignment;
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
    if (crossBetween) result.crossBetween = crossBetween as AxisCrossBetween;
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
function readAxes(plotArea: XmlElement | undefined): AxisOptions[] | undefined {
  if (!plotArea?.elements) return undefined;
  const axes: AxisOptions[] = [];
  for (const child of plotArea.elements) {
    if (!child.name) continue;
    const kind = AXIS_KIND_BY_TAG[child.name];
    if (kind) axes.push(readAxis(child, kind));
  }
  return axes.length > 0 ? axes : undefined;
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

    // CT_ChartSpace: style → clrMapOvr → pivotSource → protection → chart
    if (opts.colorMapOverride) parts.push(stringifyColorMapOverride(opts.colorMapOverride));
    if (opts.pivotSource) parts.push(stringifyPivotSource(opts.pivotSource));
    if (opts.protection) parts.push(stringifyProtection(opts.protection));

    // c:chart container
    parts.push("<c:chart>");

    // Title
    if (opts.title) {
      parts.push(stringifyTitle(opts.title));
    }

    parts.push(valEl("c:autoTitleDeleted", 0));
    // CT_Chart: autoTitleDeleted → pivotFmts → view3D
    if (opts.pivotFormats?.length) parts.push(stringifyPivotFormats(opts.pivotFormats));

    // 3D view (before plotArea per CT_Chart sequence)
    if (opts.view3D) {
      parts.push(stringifyView3D(opts.view3D));
    }

    // 3D walls and floor (CT_Chart: view3D → floor → sideWall → backWall → plotArea)
    parts.push(stringifySurface("c:floor", opts.floor));
    parts.push(stringifySurface("c:sideWall", opts.sideWall));
    parts.push(stringifySurface("c:backWall", opts.backWall));

    // c:plotArea
    parts.push(`<c:plotArea>${stringifyLayout(opts.plotAreaLayout)}`);

    // Chart type element (header + series + footer)
    parts.push(chartTypeHeader(opts));

    for (const [i, series] of opts.series.entries()) {
      parts.push(stringifySeries(i, series, opts));
    }

    parts.push(chartTypeFooter(opts));

    // Axes
    parts.push(stringifyAxes(opts));

    // Data table (CT_PlotArea: axes → dTable → spPr)
    if (opts.dataTable) parts.push(stringifyDataTable(opts.dataTable));

    parts.push("</c:plotArea>");

    // Legend
    if (opts.showLegend !== false) {
      parts.push(stringifyLegend(opts));
    }

    // CT_Chart tail: plotVisOnly → dispBlanksAs → showDLblsOverMax
    parts.push(`<c:plotVisOnly${boolVal(opts.plotVisOnly ?? true)}/>`);
    if (opts.displayBlanksAs !== undefined)
      parts.push(valEl("c:dispBlanksAs", opts.displayBlanksAs));
    if (opts.showDataLabelsOverMax !== undefined)
      parts.push(`<c:showDLblsOverMax${boolVal(opts.showDataLabelsOverMax)}/>`);

    parts.push("</c:chart>");

    // SpPr and txPr
    parts.push(noFillSpPr());
    parts.push(chartTxPr());

    // CT_ChartSpace: txPr → externalData → printSettings → (userShapes)
    if (opts.externalData) parts.push(stringifyExternalData(opts.externalData));
    if (opts.printSettings) parts.push(stringifyPrintSettings(opts.printSettings));
    // CT_ChartSpace: printSettings → userShapes → extLst
    if (opts.userShapes !== undefined)
      parts.push(`<c:userShapes r:id="${escapeXml(opts.userShapes)}"/>`);

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

    // CT_ChartSpace: clrMapOvr + protection (before c:chart)
    const colorMapOverride = readColorMapOverride(el);
    if (colorMapOverride) result.colorMapOverride = colorMapOverride;
    const protection = readProtection(el);
    if (protection) result.protection = protection;
    const pivotSource = readPivotSource(el);
    if (pivotSource) result.pivotSource = pivotSource;

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
    const pivotFormats = readPivotFormats(chart);
    if (pivotFormats) result.pivotFormats = pivotFormats;
    const view3D = readView3D(chart);
    if (view3D) result.view3D = view3D;

    // 3D walls and floor
    const floor = readSurface(chart, "c:floor");
    if (floor) result.floor = floor;
    const sideWall = readSurface(chart, "c:sideWall");
    if (sideWall) result.sideWall = sideWall;
    const backWall = readSurface(chart, "c:backWall");
    if (backWall) result.backWall = backWall;

    // Plot area — detect chart type
    const plotArea = findChild(chart, "c:plotArea");
    if (plotArea?.elements) {
      // plot-area manual layout (CT_Layout > CT_ManualLayout)
      const manualLayout = readManualLayout(findChild(plotArea, "c:layout"));
      if (manualLayout) result.plotAreaLayout = manualLayout;
      // plot-area data table (CT_PlotArea tail)
      const dataTable = readDataTable(plotArea);
      if (dataTable) result.dataTable = dataTable;
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
        readChartTypeScalars(chartTypeEl, detectedType, threeD, result);
        readChartTypeDecorations(chartTypeEl, detectedType, threeD, result);
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
              ...readSeriesCommon(serEl),
            });
          }
          result.series = bubbleSeries as BubbleSeriesData[];
        } else {
          const chartSeries: ChartSeriesData[] = [];
          let categories: string[] | undefined;

          for (const serEl of seriesEls) {
            const name = readSeriesName(serEl);

            // Read categories from first series (c:cat may be strRef/strLit/multiLvlStrRef)
            if (!categories && !result.multiLevelCategories && !result.categoryLabels) {
              const catEl = findChild(serEl, "c:cat") ?? findChild(serEl, "c:xVal");
              if (catEl) {
                const multi = readMultiLvlStrCache(catEl);
                if (multi) {
                  result.multiLevelCategories = multi;
                } else {
                  const lit = readStrLit(catEl);
                  if (lit) {
                    result.categoryLabels = lit;
                    categories = lit;
                  } else {
                    categories = readStrCache(catEl);
                  }
                }
              }
            }

            // Read values
            const valEl = findChild(serEl, "c:val") ?? findChild(serEl, "c:yVal");
            const values = valEl ? readNumCache(valEl) : [];

            chartSeries.push({ name, values, ...readSeriesCommon(serEl) });
          }

          result.series = chartSeries as ChartSeriesData[];
          if (categories?.length) result.categories = categories;
        }
      }
    }

    // Axes
    const axes = readAxes(plotArea);
    if (axes) result.axes = axes;

    // Legend — default is true, only set if explicitly absent
    const legend = findChild(chart, "c:legend");
    if (!legend) {
      result.showLegend = false;
    } else {
      const legendData = readLegend(chart);
      if (legendData?.position) result.legendPosition = legendData.position;
      if (legendData?.entries) result.legendEntries = legendData.entries;
    }

    // CT_Chart tail: plotVisOnly, dispBlanksAs, showDLblsOverMax
    const plotVisOnly = readBoolAttr(chart, "c:plotVisOnly");
    if (plotVisOnly !== undefined) result.plotVisOnly = plotVisOnly;
    const dispBlanksAs = readValStr(chart, "c:dispBlanksAs");
    if (dispBlanksAs) result.displayBlanksAs = dispBlanksAs as DisplayBlanksAs;
    const showDLblsOverMax = readBoolAttr(chart, "c:showDLblsOverMax");
    if (showDLblsOverMax !== undefined) result.showDataLabelsOverMax = showDLblsOverMax;

    // CT_ChartSpace tail: externalData (after c:txPr)
    const externalData = readExternalData(el);
    if (externalData) result.externalData = externalData;
    const printSettings = readPrintSettings(el);
    if (printSettings) result.printSettings = printSettings;
    const userShapesEl = findChild(el, "c:userShapes");
    if (userShapesEl) {
      const rid = attr(userShapesEl, "r:id");
      if (rid !== undefined) result.userShapes = rid;
    }

    return result as ChartSpaceOptions;
  },
};
