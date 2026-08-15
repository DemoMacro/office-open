/**
 * Chart type definitions — interfaces and constants.
 *
 * Class implementations have been removed; all chart XML generation
 * goes through chart-descriptors.ts (stringify/parse path).
 *
 * @module
 */

import type { ColorMappingOptions } from "../theme/theme-options";

// ── Series common (CT_Ser shared children) ──

/** Fields shared by every chart series type (name + optional decorations). */
export interface ChartSeriesCommon {
  name: string;
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
  /** 3D bar column shape (c:shape, bar3D). */
  shape?: BarShape;
  /** 3D bubble (c:bubble3D on the bubble series). */
  bubble3D?: boolean;
}

// ── BubbleSeriesData ──

export interface BubbleSeriesData extends ChartSeriesCommon {
  xValues: readonly number[];
  yValues: readonly number[];
  bubbleSize: readonly number[];
}

// ── Trendline ──

export type TrendlineType = "exp" | "linear" | "log" | "movingAvg" | "poly" | "power";

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

export type ErrorBarDirection = "both" | "x" | "y";

export type ErrorBarType = "both" | "minus" | "plus";

export type ErrorValueType = "cust" | "fixedVal" | "percentage" | "stdDev" | "stdErr";

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

export type DataLabelPosition =
  | "bestFit"
  | "b"
  | "ctr"
  | "inBase"
  | "inEnd"
  | "l"
  | "outEnd"
  | "r"
  | "t";

/** Single data-point label override (CT_DLbl). */
export interface DataLabelOptions {
  /** Series point index this label applies to (c:idx, required). */
  index: number;
  /** Drop the label for this point (c:delete). */
  delete?: boolean;
  position?: DataLabelPosition;
  /** Number format override (c:numFmt inside EG_DLblShared). */
  numberFormat?: string;
}

export interface DataLabelsOptions {
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

export interface ChartSpaceOptions {
  title?: string;
  type: ChartType;
  categories?: readonly string[];
  /** Multi-level (hierarchical) category labels (c:cat > c:multiLvlStrRef). */
  multiLevelCategories?: readonly (readonly string[])[];
  /** Literal category labels, emitted as c:strLit (c:cat > c:strLit). */
  categoryLabels?: readonly string[];
  series: readonly ChartSeriesData[] | readonly BubbleSeriesData[];
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
  /** Manual plot-area layout (c:plotArea > c:layout > c:manualLayout). */
  plotAreaLayout?: ManualLayoutOptions;
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
  highLowLines?: boolean;
  /** Up/down bars container (c:upDownBars, line/stock 2D). */
  upDownBars?: boolean;
  /** Gap width inside the up/down bars container (c:upDownBars > c:gapWidth). */
  upDownBarsGapWidth?: number;
  /** Drop lines (c:dropLines, line/area/stock). */
  dropLines?: boolean;
  /** Series lines (c:serLines, stock / 2D bar / 3D area). */
  seriesLines?: boolean;
  /** Plot-area data table (c:dTable). */
  dataTable?: DataTableOptions;
  /** ofPie variant: pie-of-pie or bar-of-pie (c:ofPieType, default "pie"). */
  ofPieType?: OfPieType;
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
  pivotFormats?: readonly PivotFormatOptions[];
  /** Surface chart color bands (c:bandFmts, after ser in surface charts). */
  bandFormats?: readonly BandFormatOptions[];
  /** Legend position (c:legendPos, defaults to "r" when legend is shown). */
  legendPosition?: LegendPosition;
  /** Legend entry overrides (c:legendEntry, inside c:legend). */
  legendEntries?: readonly LegendEntryOptions[];
  /** User-drawn shapes relationship id (c:userShapes r:id, after printSettings). */
  userShapes?: string;
}

// ── 3D view ──

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

export type AxisKind = "category" | "value" | "date" | "series";

export type AxisPosition = "b" | "l" | "r" | "t";

export type AxisTickMark = "cross" | "in" | "none" | "out";

export type AxisTickLabelPosition = "high" | "low" | "nextTo" | "none";

export type AxisCrosses = "autoZero" | "max" | "min";

export type AxisCrossBetween = "between" | "midCat";

export type AxisOrientation = "minMax" | "maxMin";

export type AxisLabelAlignment = "ctr" | "l" | "r";

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

/**
 * Per-axis configuration. Covers EG_AxShared plus kind-specific children of
 * CT_CatAx / CT_ValAx / CT_DateAx / CT_SerAx. Field order follows the XSD
 * content model; the serializer emits in that order.
 */
export interface AxisOptions {
  kind: AxisKind;
  /** c:axId (required). */
  id: number;
  /** c:crossAx target id (required). */
  crossAxisId: number;
  scaling?: AxisScalingOptions;
  delete?: boolean;
  position?: AxisPosition;
  /** c:majorGridlines presence (CT_ChartLines). */
  majorGridlines?: boolean;
  minorGridlines?: boolean;
  /** Plain-text axis title (c:title > c:rich > a:p > a:r > a:t). */
  title?: string;
  /** c:numFmt formatCode; sourceLinked defaults to 1. */
  numberFormat?: string;
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
}

// ── Series marker / data point / picture options (CT_Marker/CT_DPt/CT_PictureOptions) ──

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
  pictureOptions?: PictureOptionsOptions;
}

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

export type BarShape = "cone" | "coneToMax" | "box" | "cylinder" | "pyramid" | "pyramidToMax";

// ── Plot-area layout + 3D surfaces (CT_ManualLayout / CT_Layout / CT_Surface) ──

export type LayoutTarget = "inner" | "outer";

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

/** 3D wall/floor surface (CT_Surface thickness; spPr/pictureOptions not modeled here). */
export interface SurfaceOptions {
  /** Surface thickness — plain number or percentage string ("N%"). */
  thickness?: number | string;
}

// ── Chart-level scalars (CT_Chart tail + CT_xxxChart type-specific heads) ──

export type DisplayBlanksAs = "gap" | "span" | "zero";

export type SizeRepresents = "area" | "w";

/** Plot-area data table (CT_DTable). */
export interface DataTableOptions {
  showHorizontalBorder?: boolean;
  showVerticalBorder?: boolean;
  showOutline?: boolean;
  showLegendKeys?: boolean;
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
}

// ── Print settings (CT_PrintSettings) ──

/** Print-time header/footer text (CT_HeaderFooter). */
export interface HeaderFooterOptions {
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
export interface PageMarginsOptions {
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
  header?: number;
  footer?: number;
}

export type PageSetupOrientation = "default" | "portrait" | "landscape";

/** Print page setup (CT_PageSetup). */
export interface PageSetupOptions {
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
  headerFooter?: HeaderFooterOptions;
  pageMargins?: PageMarginsOptions;
  pageSetup?: PageSetupOptions;
  /** Legacy VML drawing relationship id (c:legacyDrawingHF r:id). */
  legacyDrawingId?: string;
}

// ── ofPie chart (bar-of-pie / pie-of-pie, CT_OfPieChart) ──

export type OfPieType = "pie" | "bar";

export type SplitType = "auto" | "cust" | "percent" | "pos" | "val";

// ── Pivot chart (CT_PivotSource / CT_PivotFmts) ──

/** Pivot chart source (CT_PivotSource: name + fmtId, both required). */
export interface PivotSourceOptions {
  name: string;
  formatId: number;
}

/** Per-series pivot format override (CT_PivotFmt). */
export interface PivotFormatOptions {
  index: number;
  marker?: MarkerOptions;
}

// ── Surface band formats (CT_BandFmts) ──

/** Surface chart color band format (CT_BandFmt: idx). */
export interface BandFormatOptions {
  index: number;
}

// ── Legend entries (CT_Legend / CT_LegendEntry) ──

export type LegendPosition = "b" | "tr" | "l" | "r" | "t";

/** Legend entry override (CT_LegendEntry: idx + delete | txPr). */
export interface LegendEntryOptions {
  index: number;
  /** Hide this legend entry (c:delete). */
  delete?: boolean;
}
