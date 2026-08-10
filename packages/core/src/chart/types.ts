/**
 * Chart type definitions — interfaces and constants.
 *
 * Class implementations have been removed; all chart XML generation
 * goes through chart-descriptors.ts (stringify/parse path).
 *
 * @module
 */

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

export const TrendlineType = {
  EXP: "exp",
  LINEAR: "linear",
  LOG: "log",
  MOVING_AVG: "movingAvg",
  POLY: "poly",
  POWER: "power",
} as const;

export type TrendlineType = (typeof TrendlineType)[keyof typeof TrendlineType];

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

export const ErrorBarDirection = {
  BOTH: "both",
  X: "x",
  Y: "y",
} as const;

export type ErrorBarDirection = (typeof ErrorBarDirection)[keyof typeof ErrorBarDirection];

export const ErrorBarType = {
  BOTH: "both",
  MINUS: "minus",
  PLUS: "plus",
} as const;

export type ErrorBarType = (typeof ErrorBarType)[keyof typeof ErrorBarType];

export const ErrorValueType = {
  CUST: "cust",
  FIXED: "fixedVal",
  PERCENTAGE: "percentage",
  STD_DEV: "stdDev",
  STD_ERR: "stdErr",
} as const;

export type ErrorValueType = (typeof ErrorValueType)[keyof typeof ErrorValueType];

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

export const DataLabelPosition = {
  BEST_FIT: "bestFit",
  B: "b",
  CTRL: "ctr",
  IN_BASE: "inBase",
  IN_END: "inEnd",
  L: "l",
  OUT_END: "outEnd",
  R: "r",
  T: "t",
} as const;

export type DataLabelPosition = (typeof DataLabelPosition)[keyof typeof DataLabelPosition];

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
  | "surface";

export type AxisChartType = Exclude<ChartType, "bubble">;

// ── ChartSpace options ──

export interface ChartSpaceOptions {
  title?: string;
  type: ChartType;
  categories?: readonly string[];
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
}

// ── 3D view ──

export const TimeUnit = {
  DAYS: "days",
  MONTHS: "months",
  YEARS: "years",
} as const;

export type TimeUnit = (typeof TimeUnit)[keyof typeof TimeUnit];

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

export const AxisKind = {
  CATEGORY: "category",
  VALUE: "value",
  DATE: "date",
  SERIES: "series",
} as const;
export type AxisKind = (typeof AxisKind)[keyof typeof AxisKind];

export const AxisPosition = {
  BOTTOM: "b",
  LEFT: "l",
  RIGHT: "r",
  TOP: "t",
} as const;
export type AxisPosition = (typeof AxisPosition)[keyof typeof AxisPosition];

export const AxisTickMark = {
  CROSS: "cross",
  IN: "in",
  NONE: "none",
  OUT: "out",
} as const;
export type AxisTickMark = (typeof AxisTickMark)[keyof typeof AxisTickMark];

export const AxisTickLabelPosition = {
  HIGH: "high",
  LOW: "low",
  NEXT_TO: "nextTo",
  NONE: "none",
} as const;
export type AxisTickLabelPosition =
  (typeof AxisTickLabelPosition)[keyof typeof AxisTickLabelPosition];

export const AxisCrosses = {
  AUTO_ZERO: "autoZero",
  MAX: "max",
  MIN: "min",
} as const;
export type AxisCrosses = (typeof AxisCrosses)[keyof typeof AxisCrosses];

export const AxisCrossBetween = {
  BETWEEN: "between",
  MID_CATEGORY: "midCat",
} as const;
export type AxisCrossBetween = (typeof AxisCrossBetween)[keyof typeof AxisCrossBetween];

export const AxisOrientation = {
  MIN_MAX: "minMax",
  MAX_MIN: "maxMin",
} as const;
export type AxisOrientation = (typeof AxisOrientation)[keyof typeof AxisOrientation];

export const AxisLabelAlignment = {
  CENTER: "ctr",
  LEFT: "l",
  RIGHT: "r",
} as const;
export type AxisLabelAlignment = (typeof AxisLabelAlignment)[keyof typeof AxisLabelAlignment];

export const BuiltInDisplayUnit = {
  HUNDREDS: "hundreds",
  THOUSANDS: "thousands",
  TEN_THOUSANDS: "tenThousands",
  HUNDRED_THOUSANDS: "hundredThousands",
  MILLIONS: "millions",
  TEN_MILLIONS: "tenMillions",
  HUNDRED_MILLIONS: "hundredMillions",
  BILLIONS: "billions",
  TRILLIONS: "trillions",
} as const;
export type BuiltInDisplayUnit = (typeof BuiltInDisplayUnit)[keyof typeof BuiltInDisplayUnit];

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

export const MarkerSymbol = {
  CIRCLE: "circle",
  DASH: "dash",
  DIAMOND: "diamond",
  DOT: "dot",
  NONE: "none",
  PICTURE: "picture",
  PLUS: "plus",
  SQUARE: "square",
  STAR: "star",
  TRIANGLE: "triangle",
  X: "x",
} as const;
export type MarkerSymbol = (typeof MarkerSymbol)[keyof typeof MarkerSymbol];

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

export const PictureFormat = {
  STACK: "stack",
  SCALE: "scale",
  STACK_SCALE: "stackScale",
  STRETCH: "stretch",
} as const;
export type PictureFormat = (typeof PictureFormat)[keyof typeof PictureFormat];

/** Picture-fill options for bar/area series (CT_PictureOptions). */
export interface PictureOptionsOptions {
  applyToFront?: boolean;
  applyToSides?: boolean;
  applyToEnd?: boolean;
  pictureFormat?: PictureFormat;
  /** Stack unit (c:pictureStackUnit val). */
  pictureStackUnit?: number;
}

export const BarShape = {
  CONE: "cone",
  CONE_TO_MAX: "coneToMax",
  BOX: "box",
  CYLINDER: "cylinder",
  PYRAMID: "pyramid",
  PYRAMID_TO_MAX: "pyramidToMax",
} as const;
export type BarShape = (typeof BarShape)[keyof typeof BarShape];

// ── Plot-area layout + 3D surfaces (CT_ManualLayout / CT_Layout / CT_Surface) ──

export const LayoutTarget = {
  INNER: "inner",
  OUTER: "outer",
} as const;
export type LayoutTarget = (typeof LayoutTarget)[keyof typeof LayoutTarget];

export const LayoutMode = {
  EDGE: "edge",
  FACTOR: "factor",
} as const;
export type LayoutMode = (typeof LayoutMode)[keyof typeof LayoutMode];

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
