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
  readonly name: string;
  readonly xValues: readonly number[];
  readonly yValues: readonly number[];
  readonly bubbleSize: readonly number[];
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
  readonly type?: TrendlineType;
  readonly name?: string;
  readonly order?: number;
  readonly period?: number;
  readonly forward?: number;
  readonly backward?: number;
  readonly intercept?: number;
  readonly dispRSqr?: boolean;
  readonly dispEq?: boolean;
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
  readonly direction?: ErrorBarDirection;
  readonly barType?: ErrorBarType;
  readonly valueType?: ErrorValueType;
  readonly value?: number;
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
  readonly position?: DataLabelPosition;
  readonly showVal?: boolean;
  readonly showCatName?: boolean;
  readonly showSerName?: boolean;
  readonly showPercent?: boolean;
  readonly showBubbleSize?: boolean;
  readonly showLeaderLines?: boolean;
}

// ── Chart series ──

export interface ChartSeriesData {
  readonly name: string;
  readonly values: readonly number[];
  readonly trendlines?: readonly TrendlineOptions[];
  readonly errorBars?: ErrorBarOptions;
  readonly dataLabels?: DataLabelsOptions;
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
  readonly title?: string;
  readonly type: ChartType;
  readonly categories?: readonly string[];
  readonly series: readonly ChartSeriesData[] | readonly BubbleSeriesData[];
  readonly showLegend?: boolean;
  readonly style?: number;
  readonly threeD?: boolean;
}

// ── 3D view ──

export const TimeUnit = {
  DAYS: "days",
  MONTHS: "months",
  YEARS: "years",
} as const;

export type TimeUnit = (typeof TimeUnit)[keyof typeof TimeUnit];

export interface View3DOptions {
  readonly rotX?: number;
  readonly rotY?: number;
  readonly depthPercent?: number;
  readonly rAngAx?: boolean;
  readonly perspective?: number;
}
