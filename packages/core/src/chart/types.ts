/**
 * Chart type definitions — interfaces and constants.
 *
 * Class implementations have been removed; all chart XML generation
 * goes through chart-descriptors.ts (stringify/parse path).
 *
 * @module
 */

// ── BubbleSeriesData ──

export interface BubbleSeriesData {
  name: string;
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

export interface DataLabelsOptions {
  position?: DataLabelPosition;
  showVal?: boolean;
  showCatName?: boolean;
  showSerName?: boolean;
  showPercent?: boolean;
  showBubbleSize?: boolean;
  showLeaderLines?: boolean;
}

// ── Chart series ──

export interface ChartSeriesData {
  name: string;
  values: readonly number[];
  trendlines?: readonly TrendlineOptions[];
  errorBars?: ErrorBarOptions;
  dataLabels?: DataLabelsOptions;
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
  rotY?: number;
  depthPercent?: number;
  rAngAx?: boolean;
  perspective?: number;
}
