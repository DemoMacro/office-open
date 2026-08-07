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
  /** Chart key for externalData reference (enables chart editing in Word) */
  chartKey?: string;
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
