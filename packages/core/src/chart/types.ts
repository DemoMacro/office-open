/**
 * Chart type definitions — interfaces and constants.
 *
 * Class implementations have been removed; all chart XML generation
 * goes through chart-descriptors.ts (stringify/parse path).
 *
 * @module
 */

import type { ShapePropertiesOptions } from "../drawing/shape-properties-desc";
import type { TextBodyOptions } from "../drawing/text/text-body";
import type { ColorMappingOptions } from "../theme/theme-options";
import type { DataType } from "../util/data-type";

// ── Series common (CT_Ser shared children) ──

/** Fields shared by every chart series type (name + optional decorations). */
export interface ChartSeriesCommon {
  /**
   * Series index across ALL groups of the chart (c:idx, unique chart-wide —
   * combo charts number secondary-group series into the same space). Omit to
   * number by array position.
   */
  index?: number;
  /**
   * Display order across the chart's series (c:order). Omit to number by
   * array position.
   */
  order?: number;
  /** Series name; absent when the source series carried no c:tx. */
  name?: string;
  /** Series name reference formula (c:tx > c:strRef > c:f) — round-trip. */
  nameFormula?: string;
  /**
   * Series name is an inline literal (c:tx > c:v, no strRef wrapper) — the
   * dominant Excel form; parse sets this when the source carried a bare c:v.
   */
  nameLiteral?: boolean;
  /** Values reference formula (c:val > c:numRef > c:f) — round-trip. */
  valueFormula?: string;
  /**
   * Values are inline literals (c:val > c:numLit instead of c:numRef) —
   * parse sets this when the source carried c:numLit.
   */
  valueLiteral?: boolean;
  /** Values number format (c:numCache > c:formatCode) — round-trip. */
  formatCode?: string;
  trendlines?: readonly TrendlineOptions[];
  errorBars?: ErrorBarOptions;
  dataLabels?: DataLabelsOptions;
  /** Per-point overrides (c:dPt), emitted after spPr in ser order. */
  dataPoints?: readonly DataPointOptions[];
  /** Line/scatter/radar marker (c:marker). */
  marker?: MarkerOptions;
  /** Invert fill for negative values (c:invertIfNegative, bar/bubble). */
  invertIfNegative?: boolean;
  /** Smooth the line (c:smooth, line/scatter). */
  smooth?: boolean;
  /** Pie slice explosion offset in percent (c:explosion). */
  explosion?: number;
  /** Picture-fill options (c:pictureOptions, bar/area). */
  pictureOptions?: PictureOptionsOptions;
  /** Series shape properties (c:ser > c:spPr: fill/outline/effects) — round-trip. */
  shapeProperties?: ShapePropertiesOptions;
  /** 3D bar column shape (c:shape, bar3D). */
  shape?: BarShape;
  /** 3D bubble (c:bubble3D on the bubble series). */
  bubble3D?: boolean;
  /**
   * Raw inner XML of the series' trailing c:extLst (CT_xxxSer tail — where
   * c16:uniqueId lives) — verbatim round-trip.
   */
  ext?: string;
}

// ── ScatterSeriesData ──

/**
 * XY numeric series (c:xVal/c:yVal) — a scatter series without bubble sizes.
 * `valueFormula`/`formatCode` describe the y reference, matching the
 * ChartSeriesData value-slot semantics.
 */
export interface ScatterSeriesData extends ChartSeriesCommon {
  xValues: readonly number[];
  yValues: readonly number[];
}

// ── BubbleSeriesData ──

export interface BubbleSeriesData extends ScatterSeriesData {
  bubbleSize: readonly number[];
}

// ── Trendline ──

/** Trendline fit; see period/order for the moving-average and polynomial degrees. */
export type TrendlineType =
  | "exponential"
  | "linear"
  | "logarithmic"
  | "movingAverage"
  | "polynomial"
  | "power";

/** Trendline label formatting (CT_TrendlineLbl — all children optional). */
export interface TrendlineLabelOptions {
  /** Number format applied to the trendline label (c:numFmt formatCode). */
  numberFormat?: string;
}

export interface TrendlineOptions {
  type?: TrendlineType;
  name?: string;
  order?: number;
  period?: number;
  forward?: number;
  backward?: number;
  intercept?: number;
  dispRSqr?: boolean;
  dispEq?: boolean;
  /** Per-trendline label formatting (emitted as c:trendlineLbl). */
  label?: TrendlineLabelOptions;
}

// ── Error bars ──

/** Which axes the error bars span: both, x only, or y only. */
export type ErrorBarDirection = "x" | "y";

/** Which side of the point gets a bar: both, minus (below/left), or plus (above/right). */
export type ErrorBarType = "both" | "minus" | "plus";

/** Error amount source: custom plus/minus values (see plusValue/minusValue), a fixed value, a percentage of the data point, standard deviation, or standard error. */
export type ErrorValueType =
  | "custom"
  | "fixedValue"
  | "percentage"
  | "standardDeviation"
  | "standardError";

export interface ErrorBarOptions {
  direction?: ErrorBarDirection;
  barType?: ErrorBarType;
  valueType?: ErrorValueType;
  value?: number;
  noEndCap?: boolean;
  /** Custom plus error value (emitted as c:plus > c:numLit). */
  plusValue?: number;
  /** Custom minus error value (emitted as c:minus > c:numLit). */
  minusValue?: number;
}

// ── Data labels ──

/** Data-label placement; "bestFit" lets the renderer choose (pie). */
export type DataLabelPosition =
  | "bestFit"
  | "bottom"
  | "center"
  | "insideBase"
  | "insideEnd"
  | "left"
  | "outsideEnd"
  | "right"
  | "top";

/** Single data-point label override (CT_DLbl). */
export interface DataLabelOptions {
  /** Series point index this label applies to (c:idx, required). */
  index: number;
  /** Drop the label for this point (c:delete). */
  delete?: boolean;
  /**
   * Custom label text (c:tx > c:rich) — replaces the auto-generated value
   * label with authored rich text.
   */
  text?: TextBodyOptions;
  /**
   * Manual label placement (c:layout > c:manualLayout, before EG_DLblShared).
   * true = the bare `<c:layout/>` form sources carry.
   */
  layout?: ManualLayoutOptions | true;
  position?: DataLabelPosition;
  /** Number format override (c:numFmt inside EG_DLblShared). */
  numberFormat?: string;
  /** Label fill/outline (c:dLbl > c:spPr) — round-trip. */
  shapeProperties?: ShapePropertiesOptions;
  /** Label text formatting (c:dLbl > c:txPr) — round-trip. */
  textProperties?: TextBodyOptions;
  showLegendKey?: boolean;
  showVal?: boolean;
  showCatName?: boolean;
  showSerName?: boolean;
  showPercent?: boolean;
  showBubbleSize?: boolean;
  separator?: string;
}

export interface DataLabelsOptions {
  /**
   * Labels deleted at the group level (c:delete inside c:dLbls) — Office
   * writes val="0" explicitly to keep labels on; round-trips the attribute.
   */
  delete?: boolean;
  /** Label number format (c:numFmt formatCode, before the position). */
  numberFormat?: string;
  /** Shared label fill/outline (c:dLbls > c:spPr) — round-trip. */
  shapeProperties?: ShapePropertiesOptions;
  /** Shared label text formatting (c:dLbls > c:txPr) — round-trip. */
  textProperties?: TextBodyOptions;
  position?: DataLabelPosition;
  showVal?: boolean;
  showCatName?: boolean;
  showSerName?: boolean;
  showPercent?: boolean;
  showBubbleSize?: boolean;
  showLegendKey?: boolean;
  showLeaderLines?: boolean;
  separator?: string;
  /** Per-point label overrides (c:dLbl, emitted before the shared settings). */
  labels?: readonly DataLabelOptions[];
  /** Emit a c:leaderLines element for default-styled leader lines. */
  leaderLines?: boolean;
}

// ── Chart series ──

export interface ChartSeriesData extends ChartSeriesCommon {
  values: readonly number[];
}

// ── Chart types ──

/**
 * Chart family; drives the c:*Chart plot element. "ofPie" is the
 * pie-of-pie / bar-of-pie family (see ofPieType).
 */
export type ChartType =
  | "column"
  | "bar"
  | "line"
  | "pie"
  | "area"
  | "scatter"
  | "bubble"
  | "doughnut"
  | "radar"
  | "stock"
  | "surface"
  | "ofPie";

// ── ChartSpace options ──

/**
 * Chart title (CT_Title). A plain string emits the simple c:rich form; an
 * empty string the bare `<c:title/>` placeholder; the object form round-trips
 * layout/overlay/spPr/txPr decorations alongside the text.
 */
export interface ChartTitleOptions {
  /**
   * Title text — a plain string emits a single default run; a TextBodyOptions
   * carries the full c:rich body (paragraph alignment, run properties).
   */
  text?: string | TextBodyOptions;
  /** Title layout (c:layout) — `true` emits the bare `<c:layout/>` Office
   *  writes; an object is the manual layout. */
  layout?: boolean | ManualLayoutOptions;
  /** Title overlaps the plot (c:overlay). */
  overlay?: boolean;
  shapeProperties?: ShapePropertiesOptions;
  textProperties?: TextBodyOptions;
}

/**
 * Word 2010+ chart style form: c14:style in mc:Choice plus the equivalent
 * c:style in mc:Fallback.
 */
export interface ChartStyle2010Options {
  /** c14:style/@val (the 101+ style ids). */
  style: number;
  /** mc:Fallback c:style/@val. */
  fallbackStyle?: number;
}

/** Chart lines with formatting (CT_ChartLines: spPr only). */
export interface ChartLinesOptions {
  /** Line fill/outline (spPr). */
  shapeProperties?: ShapePropertiesOptions;
}

/** Series grouping mode (c:grouping val) — bar/column/line/area groups. */
export type ChartGrouping = "clustered" | "standard" | "stacked" | "percentStacked";

/**
 * Additional chart groups in a combo chart (CT_PlotArea may carry several
 * *Chart groups; each secondary one shares the category source and the axes
 * list with the main group).
 */
export interface SecondaryChartGroupOptions {
  /** Chart type of this group (drives the c:*Chart element and its header). */
  type: ChartType;
  /** This group's series; categories come from the chart-level source. */
  series: readonly ChartSeriesData[];
  /** Series grouping (c:grouping); defaults per type like the main group. */
  grouping?: ChartGrouping;
  /** Vary data-point colors (group-level c:varyColors). */
  varyColors?: boolean;
  /** Show line-chart markers (group-level c:marker, line 2D). */
  markers?: boolean;
  /** Smooth lines (group-level c:smooth, line 2D). */
  smooth?: boolean;
  /** Group-level data labels (c:dLbls after the ser elements). */
  dataLabels?: DataLabelsOptions;
  /** Bar/column gap width as percent (c:gapWidth, bar group). */
  gapWidth?: number;
  /** Bar/column overlap as percent (c:overlap, 2D bar group). */
  overlap?: number;
  /** 3D gap depth as percent (c:gapDepth). */
  gapDepth?: number;
  /** Series connector lines (c:serLines, bar/ofPie groups). */
  seriesLines?: boolean | ChartLinesOptions;
  /** Drop lines (c:dropLines, line/area). */
  dropLines?: boolean | ChartLinesOptions;
  /** High-low lines (c:hiLowLines, line/stock). */
  highLowLines?: boolean | ChartLinesOptions;
  /** Up/down bars container (c:upDownBars, line). */
  upDownBars?: boolean;
  /** Gap width inside the up/down bars container (c:upDownBars > c:gapWidth). */
  upDownBarsGapWidth?: number;
  /** axId pair referencing entries of the shared axes list. */
  axisIds?: readonly number[];
}

export interface ChartSpaceOptions {
  /**
   * Namespace dialect of the chart part. A strict (ISO/IEC 29500 Strict,
   * purl.oclc.org) source package rejects a chart re-serialized with
   * transitional namespaces — parse records the dialect from the root element
   * and stringify re-declares the matching namespace set.
   */
  dialect?: "transitional" | "strict";
  /**
   * Chart title (c:title). An empty string round-trips a title placeholder
   * that carries no text (bare `<c:title/>`, the legacy-Word auto-title form).
   */
  title?: string | ChartTitleOptions;
  /** 1904 date system (c:date1904) — emitted only when the source had it. */
  date1904?: boolean;
  /** Chart UI language (c:lang, e.g. "en-US") — emitted only when set. */
  lang?: string;
  /** Rounded chart-area corners (c:roundedCorners) — emitted only when set. */
  roundedCorners?: boolean;

  /**
   * Word 2010+ chart style — c14:style carried via mc:AlternateContent with a
   * c:style fallback for older readers. Takes precedence over `style`.
   */
  style2010?: ChartStyle2010Options;
  /** Auto-generated title suppressed (c:autoTitleDeleted) — emitted only when set. */
  autoTitleDeleted?: boolean;
  /** Chart-area shape properties (c:spPr after c:chart) — round-trip. */
  shapeProperties?: ShapePropertiesOptions;
  /** Chart-space default text (c:txPr, a CT_TextBody) — round-trip. */
  textProperties?: TextBodyOptions;
  type: ChartType;
  categories?: readonly string[];
  /**
   * Categories are numeric (c:cat carries c:numRef/c:numCache instead of
   * c:strRef/c:strCache). Numeric categories render on a value-formatted
   * axis; parse sets this when the source used numRef.
   */
  numericCategories?: boolean;
  /** Category reference formula (c:cat > c:*Ref > c:f) — round-trip. */
  categoryFormula?: string;
  /** Category number format (c:numCache > c:formatCode, numeric categories). */
  categoryFormatCode?: string;
  /** Multi-level (hierarchical) category labels (c:cat > c:multiLvlStrRef). */
  multiLevelCategories?: readonly (readonly string[])[];
  /** Literal category labels, emitted as c:strLit (c:cat > c:strLit). */
  categoryLabels?: readonly string[];
  // Bubble stays an explicit union arm: TS would accept a BubbleSeriesData[]
  // as ScatterSeriesData[] (structural subtype), but the generated schema's
  // closed-world properties reject the extra bubbleSize field.
  series: readonly ChartSeriesData[] | readonly ScatterSeriesData[] | readonly BubbleSeriesData[];
  /**
   * Vary data-point colors (chart-group-level c:varyColors). Emitted only
   * when set — the corpus is mixed and fresh output stays bare.
   */
  varyColors?: boolean;
  /**
   * Series grouping (c:grouping val). Absent → the per-type default the
   * header always wrote (clustered for bar/column, standard for line/area).
   */
  grouping?: ChartGrouping;
  /**
   * Show line-chart markers (c:marker CT_Boolean on c:lineChart, 2D only).
   * Emitted only when set.
   */
  markers?: boolean;
  /**
   * Smooth lines for the whole line-chart group (c:smooth CT_Boolean on
   * c:lineChart, after c:marker) — distinct from the per-series `smooth`.
   * Emitted only when set.
   */
  smooth?: boolean;
  /** 3D bar column shape (chart-group-level c:shape on c:bar3DChart). */
  shape?: BarShape;
  /** Chart-group-level data labels (c:dLbls after the ser elements). */
  dataLabels?: DataLabelsOptions;
  showLegend?: boolean;
  style?: number;
  threeD?: boolean;
  view3D?: View3DOptions;
  /**
   * Per-axis configuration. Omitted, the library emits sensible default axes
   * derived from chart type; provided, each axis is serialized verbatim and
   * parsed back, enabling full round-trip of gridlines/units/tick marks/etc.
   */
  axes?: readonly AxisOptions[];
  /**
   * Axis references (c:axId sequence on the chart group) — round-trip.
   * Normally derived from `axes`; kept when the source sequence differs
   * (legacy Word writes a dangling `axId val="0"` for a missing third axis).
   */
  axisIds?: readonly number[];
  /**
   * Second chart group in a combo chart (e.g. lines over bars on two axes).
   * Shares the plot area, categories, and the axes list; carries its own
   * series and group-level flags.
   */
  secondaryGroups?: readonly SecondaryChartGroupOptions[];
  /** Manual plot-area layout (c:plotArea > c:layout > c:manualLayout). */
  plotAreaLayout?: ManualLayoutOptions;
  /** Plot-area fill and outline (c:plotArea's trailing c:spPr) — round-trip. */
  plotAreaShapeProperties?: ShapePropertiesOptions;
  /** 3D floor (c:floor). */
  floor?: SurfaceOptions;
  /** 3D side wall (c:sideWall). */
  sideWall?: SurfaceOptions;
  /** 3D back wall (c:backWall). */
  backWall?: SurfaceOptions;
  /** Plot visible cells only (c:plotVisOnly, defaults to true when omitted). */
  plotVisOnly?: boolean;
  /** How blank cells are plotted (c:dispBlanksAs). */
  displayBlanksAs?: DisplayBlanksAs;
  /** Show data labels over the chart maximum (c:showDLblsOverMax). */
  showDataLabelsOverMax?: boolean;
  /** Bar/column gap width as percent of width (c:gapWidth). */
  gapWidth?: number;
  /** Bar/column overlap as percent (c:overlap, 2D only). */
  overlap?: number;
  /** 3D bar/column gap depth as percent (c:gapDepth). */
  gapDepth?: number;
  /** Pie/doughnut first slice angle in degrees (c:firstSliceAng). */
  firstSliceAngle?: number;
  /** Pie/doughnut doughnut hole size as percent (c:holeSize). */
  holeSize?: number;
  /** Bubble scale as percent of default size (c:bubbleScale). */
  bubbleScale?: number;
  /** Show negative-value bubbles (c:showNegBubbles). */
  showNegativeBubbles?: boolean;
  /** What the bubble size represents (c:sizeRepresents). */
  sizeRepresents?: SizeRepresents;
  /** Surface chart wireframe rendering (c:wireframe). */
  wireframe?: boolean;
  /** High-low lines (c:hiLowLines, line/stock 2D). */
  highLowLines?: boolean | ChartLinesOptions;
  /** Up/down bars container (c:upDownBars, line/stock 2D). */
  upDownBars?: boolean;
  /** Gap width inside the up/down bars container (c:upDownBars > c:gapWidth). */
  upDownBarsGapWidth?: number;
  /** Drop lines (c:dropLines, line/area/stock). */
  dropLines?: boolean | ChartLinesOptions;
  /** Series lines (c:serLines, stock / 2D bar / 3D area). */
  seriesLines?: boolean | ChartLinesOptions;
  /** Plot-area data table (c:dTable). */
  dataTable?: DataTableOptions;
  /** ofPie variant: pie-of-pie or bar-of-pie (c:ofPieType, default "pie"). */
  ofPieType?: OfPieType;
  /** Radar variant (c:radarStyle, default "standard"). */
  radarStyle?: RadarStyle;
  /** ofPie split behavior (c:splitType, default "auto"). */
  splitType?: SplitType;
  /** ofPie split position (c:splitPos, used with splitType pos/val). */
  splitPosition?: number;
  /** ofPie custom split point indices (c:custSplit > c:secondPiePt, with splitType "cust"). */
  customSplitPoints?: readonly number[];
  /** ofPie second-pie size — number (5–200) or percent string (c:secondPieSize). */
  secondPieSize?: number | string;
  /** Theme color mapping override (c:clrMapOvr, before c:chart). */
  colorMappingOverride?: Partial<ColorMappingOptions>;
  /** Chart protection (c:protection, before c:chart). */
  protection?: ProtectionOptions;
  /** External linked workbook (c:externalData, after c:txPr). */
  externalData?: ExternalDataOptions;
  /** Print settings (c:printSettings, after c:externalData). */
  printSettings?: PrintSettingsOptions;
  /** Pivot chart source (c:pivotSource, after clrMapOvr, before protection). */
  pivotSource?: PivotSourceOptions;
  /** Pivot chart per-series formats (c:pivotFmts, after autoTitleDeleted). */
  pivotFormats?: readonly ChartPivotFormatOptions[];
  /** Surface chart color bands (c:bandFmts, after ser in surface charts). */
  bandFormats?: readonly BandFormatOptions[];
  /** Legend position (c:legendPos, defaults to "r" when legend is shown). */
  legendPosition?: LegendPosition;
  /** Legend overlaps the plot (c:overlay). */
  legendOverlay?: boolean;
  /** Legend fill and outline (c:legend's c:spPr) — round-trip. */
  legendShapeProperties?: ShapePropertiesOptions;
  /**
   * Legend manual layout (c:legend's c:layout) — true emits the bare
   * `<c:layout/>` form, an object the manual layout.
   */
  legendLayout?: boolean | ManualLayoutOptions;
  /** Legend text (c:legend's c:txPr, a CT_TextBody) — round-trip. */
  legendTextProperties?: TextBodyOptions;
  /** Legend entry overrides (c:legendEntry, inside c:legend). */
  legendEntries?: readonly LegendEntryOptions[];
  /** User-drawn shapes relationship id (c:userShapes r:id, after printSettings). */
  userShapes?: string;
  /**
   * Raw inner XML of the chart-space trailing c:extLst (after userShapes —
   * c14:pivotOptions' home). Round-trip only, same contract as series ext.
   */
  ext?: string;
}

// ── 3D view ──

/** Time-axis unit granularity (ST_TimeUnit). */
export type TimeUnit = "days" | "months" | "years";

export interface View3DOptions {
  rotX?: number;
  /** Height percentage of the chart (XSD CT_HPercent). */
  hPercent?: number;
  rotY?: number;
  depthPercent?: number;
  rAngAx?: boolean;
  perspective?: number;
}

// ── Chart axes (EG_AxShared + CT_CatAx / CT_ValAx / CT_DateAx / CT_SerAx) ──

/** Axis role; decides the CT_CatAx/CT_ValAx/CT_DateAx/CT_SerAx element. */
export type AxisKind = "category" | "value" | "date" | "series";

/** Axis side (c:axPos). */
export type AxisPosition = "bottom" | "left" | "right" | "top";

/** Tick-mark style: "cross" spanning the axis, "in" inside the plot, "out" outside, "none". */
export type AxisTickMark = "cross" | "in" | "none" | "out";

/** Tick-label placement: "high" top of the plot, "low" bottom, "nextTo" beside the axis, "none". */
export type AxisTickLabelPosition = "high" | "low" | "nextTo" | "none";

/** Where the crossing axis meets this one: at zero, the maximum, or the minimum. */
export type AxisCrosses = "zero" | "max" | "min";

/** Value-axis crossing mode: between category bands or at the band middle. */
export type AxisCrossBetween = "between" | "middleOfCategory";

/** Value direction along the axis. */
export type AxisOrientation = "ascending" | "descending";

/** Category tick-label alignment. */
export type AxisLabelAlignment = "center" | "left" | "right";

/** Built-in value-axis scaling unit (ST_BuiltInUnit), e.g. "millions" divides tick values by 1e6. */
export type BuiltInDisplayUnit =
  | "hundreds"
  | "thousands"
  | "tenThousands"
  | "hundredThousands"
  | "millions"
  | "tenMillions"
  | "hundredMillions"
  | "billions"
  | "trillions";

/** Axis scale bounds and orientation (CT_Scaling). */
export interface AxisScalingOptions {
  /** Logarithmic base, 2-1000 (c:logBase val, required when present). */
  logBase?: number;
  orientation?: AxisOrientation;
  max?: number;
  min?: number;
}

/** Display-unit scaling for a value axis (CT_DispUnits). */
export interface DisplayUnitsOptions {
  /** Custom display-unit value (c:custUnit); mutually exclusive with builtInUnit. */
  customUnit?: number;
  /** Built-in display unit (c:builtInUnit). */
  builtInUnit?: BuiltInDisplayUnit;
  /** Emit c:dispUnitsLbl (presence only; CT_DispUnitsLbl body not modeled here). */
  label?: boolean;
}

/** Explicit axis c:numFmt (CT_NumFmt) with a decoupled sourceLinked flag. */
export interface AxisNumberFormatOptions {
  formatCode: string;
  /** 1 (default) links the format to the source data; 0 decouples it. */
  sourceLinked?: boolean;
}

/**
 * Per-axis configuration. Covers EG_AxShared plus kind-specific children of
 * CT_CatAx / CT_ValAx / CT_DateAx / CT_SerAx. Field order follows the XSD
 * content model; the serializer emits in that order.
 */
export interface AxisOptions {
  kind: AxisKind;
  /**
   * c:axId. Pure internal wiring — the axis pair references each other, so a
   * fresh document may omit it and inherit the default axis slot's id.
   */
  id?: number;
  /** c:crossAx target id; defaults to the default axis slot's pairing. */
  crossAxisId?: number;
  scaling?: AxisScalingOptions;
  delete?: boolean;
  position?: AxisPosition;
  /**
   * c:majorGridlines (CT_ChartLines) — true emits the bare element, an
   * object carries its spPr/txPr decorations.
   */
  majorGridlines?: boolean | ChartLinesOptions;
  minorGridlines?: boolean | ChartLinesOptions;
  /** Axis title (c:title) — plain text or the full CT_Title object form. */
  title?: string | ChartTitleOptions;
  /**
   * c:numFmt — a plain string is the formatCode with the default
   * sourceLinked=1; the object form carries an explicit sourceLinked (a
   * source that decoupled the axis format from the data writes "0").
   */
  numberFormat?: string | AxisNumberFormatOptions;
  majorTickMark?: AxisTickMark;
  minorTickMark?: AxisTickMark;
  tickLabelPosition?: AxisTickLabelPosition;
  /** Mutually exclusive with crossesAt (XSD choice). */
  crosses?: AxisCrosses;
  crossesAt?: number;
  // CT_CatAx
  auto?: boolean;
  labelAlignment?: AxisLabelAlignment;
  /** c:lblOffset — plain number or percentage string ("N%"). */
  labelOffset?: number | string;
  tickLabelSkip?: number;
  tickMarkSkip?: number;
  noMultiLevelLabel?: boolean;
  // CT_DateAx
  baseTimeUnit?: TimeUnit;
  majorTimeUnit?: TimeUnit;
  minorTimeUnit?: TimeUnit;
  // CT_ValAx + CT_DateAx
  majorUnit?: number;
  minorUnit?: number;
  // CT_ValAx
  crossBetween?: AxisCrossBetween;
  displayUnits?: DisplayUnitsOptions;
  /** Axis shape properties (c:spPr between tickLblPos and crossAx) — round-trip. */
  shapeProperties?: ShapePropertiesOptions;
  /** Axis text (c:txPr, a CT_TextBody) — round-trip. */
  textProperties?: TextBodyOptions;
}

// ── Series marker / data point / picture options (CT_Marker/CT_DPt/CT_PictureOptions) ──

/** Point-marker glyph (ST_MarkerStyle). */
export type MarkerSymbol =
  | "circle"
  | "dash"
  | "diamond"
  | "dot"
  | "none"
  | "picture"
  | "plus"
  | "square"
  | "star"
  | "triangle"
  | "x";

/** Line/scatter/radar point marker (CT_Marker). */
export interface MarkerOptions {
  symbol?: MarkerSymbol;
  /** Marker size, 2-72 (c:size val). */
  size?: number;
  /** Marker fill/outline (c:marker > c:spPr) — round-trip. */
  shapeProperties?: ShapePropertiesOptions;
}

/** Per-point override (CT_DPt). Field order follows the XSD content model. */
export interface DataPointOptions {
  /** Point index (c:idx, required). */
  index: number;
  invertIfNegative?: boolean;
  marker?: MarkerOptions;
  /** 3D bubble on this point (c:bubble3D). */
  bubble3D?: boolean;
  /** Pie explosion offset for this point (c:explosion). */
  explosion?: number;
  /** Per-point fill/outline (c:dPt > c:spPr) — round-trip. */
  shapeProperties?: ShapePropertiesOptions;
  pictureOptions?: PictureOptionsOptions;
}

/** Picture-fill mode (ST_PictureFormat): "stack" tile at natural size, "scale" stretch to fit, "stackScale" tile scaled by pictureStackUnit, "stretch". */
export type PictureFormat = "stack" | "scale" | "stackScale" | "stretch";

/** Picture-fill options for bar/area series (CT_PictureOptions). */
export interface PictureOptionsOptions {
  applyToFront?: boolean;
  applyToSides?: boolean;
  applyToEnd?: boolean;
  pictureFormat?: PictureFormat;
  /** Stack unit (c:pictureStackUnit val). */
  pictureStackUnit?: number;
}

/** 3D bar column shape (ST_Shape): "coneToMax"/"pyramidToMax" taper to a point at the full height; plain "cone"/"pyramid" taper halfway, then run cylindrically. */
export type BarShape = "cone" | "coneToMax" | "box" | "cylinder" | "pyramid" | "pyramidToMax";

// ── Plot-area layout + 3D surfaces (CT_ManualLayout / CT_Layout / CT_Surface) ──

/** Manual-layout region (ST_LayoutTarget): "inner" plot area only, "outer" including tick labels and axis titles. */
export type LayoutTarget = "inner" | "outer";

/** How x/y/w/h are read (ST_LayoutMode): "edge" offset from the chart edge, "factor" fraction of the plot area. */
export type LayoutMode = "edge" | "factor";

/** Manual plot-area layout (CT_ManualLayout). Field order follows the XSD. */
export interface ManualLayoutOptions {
  layoutTarget?: LayoutTarget;
  xMode?: LayoutMode;
  yMode?: LayoutMode;
  wMode?: LayoutMode;
  hMode?: LayoutMode;
  x?: number;
  y?: number;
  w?: number;
  h?: number;
}

/** 3D wall/floor surface (CT_Surface: thickness → spPr → pictureOptions). */
export interface SurfaceOptions {
  /** Surface thickness — plain number or percentage string ("N%"). */
  thickness?: number | string;
  /** Wall/floor fill and outline (c:spPr). */
  shapeProperties?: ShapePropertiesOptions;
}

// ── Chart-level scalars (CT_Chart tail + CT_xxxChart type-specific heads) ──

/** How blank cells plot (ST_DispBlanksAs): "gap" skip with a line break, "span" bridge with a straight line, "zero" treat as 0. */
export type DisplayBlanksAs = "gap" | "span" | "zero";

/** What the bubble size drives: the area or the width. */
export type SizeRepresents = "area" | "width";

/** Plot-area data table (CT_DTable). */
export interface DataTableOptions {
  showHorizontalBorder?: boolean;
  showVerticalBorder?: boolean;
  showOutline?: boolean;
  showLegendKeys?: boolean;
  /** Table line/fill (c:spPr). */
  shapeProperties?: ShapePropertiesOptions;
  /** Table text (c:txPr). */
  textProperties?: TextBodyOptions;
}

// ── ChartSpace-level containers (CT_ChartSpace) ──

/** Chart protection flags (CT_Protection). */
export interface ProtectionOptions {
  chartObject?: boolean;
  data?: boolean;
  formatting?: boolean;
  selection?: boolean;
  userInterface?: boolean;
}

/** External linked workbook reference (CT_ExternalData). */
export interface ExternalDataOptions {
  relationshipId: string;
  autoUpdate?: boolean;
  /**
   * Embedded workbook bytes (round-trip). Word links each chart to an embedded
   * xlsx via the chart part's own rels; without the bytes the c:externalData
   * rId dangles and "Edit Data" breaks. Omit for fresh authoring.
   */
  data?: DataType;
  /** Source embeddings file basename, e.g. "Microsoft_Excel_Worksheet1.xlsx". */
  fileName?: string;
}

// ── Print settings (CT_PrintSettings) ──

/** Print-time header/footer text (CT_HeaderFooter). */
export interface ChartHeaderFooterOptions {
  oddHeader?: string;
  oddFooter?: string;
  evenHeader?: string;
  evenFooter?: string;
  firstHeader?: string;
  firstFooter?: string;
  alignWithMargins?: boolean;
  differentOddEven?: boolean;
  differentFirst?: boolean;
}

/** Print page margins in inches (CT_PageMargins, all XSD-required). */
export interface ChartPageMarginsOptions {
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
  header?: number;
  footer?: number;
}

/** Print orientation (ST_Orientation): "default" keeps the printer's own setting. */
export type PageSetupOrientation = "default" | "portrait" | "landscape";

/** Print page setup (CT_PageSetup). */
export interface ChartPageSetupOptions {
  paperSize?: number;
  /** Paper height as a UniversalMeasure string (mm/cm/in/...). */
  paperHeight?: string;
  /** Paper width as a UniversalMeasure string. */
  paperWidth?: string;
  firstPageNumber?: number;
  orientation?: PageSetupOrientation;
  blackAndWhite?: boolean;
  draft?: boolean;
  useFirstPageNumber?: boolean;
  horizontalDpi?: number;
  verticalDpi?: number;
  copies?: number;
}

/** Print settings (CT_PrintSettings). */
export interface PrintSettingsOptions {
  headerFooter?: ChartHeaderFooterOptions;
  pageMargins?: ChartPageMarginsOptions;
  pageSetup?: ChartPageSetupOptions;
  /** Legacy VML drawing relationship id (c:legacyDrawingHF r:id). */
  legacyDrawingId?: string;
}

// ── ofPie chart (bar-of-pie / pie-of-pie, CT_OfPieChart) ──

/** ofPie variant (ST_OfPieType): "pie" pie-of-pie, "bar" bar-of-pie. */
export type OfPieType = "pie" | "bar";

// ── Radar chart (CT_RadarChart) ──

/** Radar variant (ST_RadarStyle): "standard" lines, "marker" lines with point markers, "filled" polygon. */
export type RadarStyle = "standard" | "marker" | "filled";

/**
 * ofPie split rule: "auto" lets Excel decide, "custom" takes the indices in
 * customSplitPoints, "percent" the smallest values under splitPosition %,
 * "position" the last splitPosition points, "value" points below splitPosition.
 */
export type SplitType = "auto" | "custom" | "percent" | "position" | "value";

// ── Pivot chart (CT_PivotSource / CT_PivotFmts) ──

/** Pivot chart source (CT_PivotSource: name + fmtId, both required). */
export interface PivotSourceOptions {
  name: string;
  formatId: number;
}

/** Per-series pivot format override (CT_PivotFmt). */
export interface ChartPivotFormatOptions {
  index: number;
  marker?: MarkerOptions;
}

// ── Surface band formats (CT_BandFmts) ──

/** Surface chart color band format (CT_BandFmt: idx). */
export interface BandFormatOptions {
  index: number;
}

// ── Legend entries (CT_Legend / CT_LegendEntry) ──

/** Legend placement (c:legendPos); "topRight" is the corner slot. */
export type LegendPosition = "bottom" | "topRight" | "left" | "right" | "top";

/** Legend entry override (CT_LegendEntry: idx + delete | txPr). */
export interface LegendEntryOptions {
  index: number;
  /** Hide this legend entry (c:delete). */
  delete?: boolean;
  /** Per-entry text properties (c:legendEntry/c:txPr) — the non-delete form. */
  textProperties?: TextBodyOptions;
}
