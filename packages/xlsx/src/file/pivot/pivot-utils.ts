/**
 * Pivot table utility types and helper functions.
 *
 * @module
 */

/** Aggregation function for data fields (maps to ST_DataConsolidateFunction). */
export const ConsolidateFunction = {
  SUM: "sum",
  COUNT: "count",
  AVERAGE: "average",
  MAX: "max",
  MIN: "min",
  PRODUCT: "product",
  COUNT_NUMS: "countNums",
  STD_DEV: "stdDev",
  STD_DEV_P: "stdDevp",
  VAR: "var",
  VAR_P: "varp",
} as const;

export type ConsolidateFunction = (typeof ConsolidateFunction)[keyof typeof ConsolidateFunction];

/** A data field definition for pivot table aggregation. */
export interface PivotDataField {
  /** Source column name to aggregate */
  readonly field: string;
  /** Aggregation function (default: "sum") */
  readonly summarize?: ConsolidateFunction;
  /** Custom name for the data field (default: "Sum of {field}") */
  readonly name?: string;
}

/** Options for a single pivot table on a worksheet. */
export interface PivotTableOptions {
  /** Pivot table name (default: "PivotTable{N}") */
  readonly name?: string;
  /** Source data range, e.g. "A1:D11" — must be on the same or a different sheet */
  readonly source: string;
  /** Source sheet name (default: current sheet) */
  readonly sourceSheet?: string;
  /** Target cell for the pivot table output (default: "A3") */
  readonly location?: string;
  /** Field names to use as row labels */
  readonly rows: readonly string[];
  /** Field names to use as column labels */
  readonly columns?: readonly string[];
  /** Data fields with aggregation settings */
  readonly data: readonly PivotDataField[];
  /** Pivot style name (default: "PivotStyleLight16") */
  readonly style?: string;
}

/** Parsed source data for pivot cache generation. */
export interface PivotSourceData {
  readonly fieldNames: readonly string[];
  readonly records: readonly (string | number)[][];
}

/**
 * Extract unique values from source data for a given field index.
 */
export function collectUniqueValues(
  records: readonly (string | number)[][],
  fieldIdx: number,
): (string | number)[] {
  const seen = new Set<string>();
  const result: (string | number)[] = [];
  for (const row of records) {
    const val = row[fieldIdx];
    const key = String(val);
    if (!seen.has(key)) {
      seen.add(key);
      result.push(val);
    }
  }
  return result;
}

/**
 * Check if a field is numeric (all non-empty values are numbers).
 */
export function isNumericField(records: readonly (string | number)[][], fieldIdx: number): boolean {
  for (const row of records) {
    const val = row[fieldIdx];
    if (typeof val === "string" && val !== "") return false;
  }
  return true;
}

/**
 * Aggregate values using the specified function.
 */
export function aggregate(values: readonly number[], func: ConsolidateFunction): number {
  if (values.length === 0) return 0;
  switch (func) {
    case "sum":
      return values.reduce((a, b) => a + b, 0);
    case "count":
    case "countNums":
      return values.length;
    case "average":
      return values.reduce((a, b) => a + b, 0) / values.length;
    case "max":
      return Math.max(...values);
    case "min":
      return Math.min(...values);
    case "product":
      return values.reduce((a, b) => a * b, 1);
    case "var":
      return sampleVariance(values);
    case "varp":
      return populationVariance(values);
    case "stdDev":
      return Math.sqrt(sampleVariance(values));
    case "stdDevp":
      return Math.sqrt(populationVariance(values));
    default:
      return values.reduce((a, b) => a + b, 0);
  }
}

function populationVariance(values: readonly number[]): number {
  const mean = values.reduce((a, b) => a + b, 0) / values.length;
  return values.reduce((sum, v) => sum + (v - mean) ** 2, 0) / values.length;
}

function sampleVariance(values: readonly number[]): number {
  if (values.length < 2) return 0;
  const mean = values.reduce((a, b) => a + b, 0) / values.length;
  return values.reduce((sum, v) => sum + (v - mean) ** 2, 0) / (values.length - 1);
}
