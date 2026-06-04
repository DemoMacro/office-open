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

/** Pivot filter type (ST_PivotFilterType) */
export const PivotFilterType = {
  UNKNOWN: "unknown",
  COUNT: "count",
  PERCENT: "percent",
  SUM: "sum",
  CAPTION_EQUAL: "captionEqual",
  CAPTION_NOT_EQUAL: "captionNotEqual",
  CAPTION_BEGINS_WITH: "captionBeginsWith",
  CAPTION_NOT_BEGINS_WITH: "captionNotBeginsWith",
  CAPTION_ENDS_WITH: "captionEndsWith",
  CAPTION_NOT_ENDS_WITH: "captionNotEndsWith",
  CAPTION_CONTAINS: "captionContains",
  CAPTION_NOT_CONTAINS: "captionNotContains",
  CAPTION_GREATER_THAN: "captionGreaterThan",
  CAPTION_GREATER_THAN_OR_EQUAL: "captionGreaterThanOrEqual",
  CAPTION_LESS_THAN: "captionLessThan",
  CAPTION_LESS_THAN_OR_EQUAL: "captionLessThanOrEqual",
  CAPTION_BETWEEN: "captionBetween",
  CAPTION_NOT_BETWEEN: "captionNotBetween",
  VALUE_EQUAL: "valueEqual",
  VALUE_NOT_EQUAL: "valueNotEqual",
  VALUE_GREATER_THAN: "valueGreaterThan",
  VALUE_GREATER_THAN_OR_EQUAL: "valueGreaterThanOrEqual",
  VALUE_LESS_THAN: "valueLessThan",
  VALUE_LESS_THAN_OR_EQUAL: "valueLessThanOrEqual",
  VALUE_BETWEEN: "valueBetween",
  VALUE_NOT_BETWEEN: "valueNotBetween",
  DATE_EQUAL: "dateEqual",
  DATE_NOT_EQUAL: "dateNotEqual",
  DATE_OLDER_THAN: "dateOlderThan",
  DATE_OLDER_THAN_OR_EQUAL: "dateOlderThanOrEqual",
  DATE_NEWER_THAN: "dateNewerThan",
  DATE_NEWER_THAN_OR_EQUAL: "dateNewerThanOrEqual",
  DATE_BETWEEN: "dateBetween",
  DATE_NOT_BETWEEN: "dateNotBetween",
  TOMORROW: "tomorrow",
  TODAY: "today",
  YESTERDAY: "yesterday",
  NEXT_WEEK: "nextWeek",
  THIS_WEEK: "thisWeek",
  LAST_WEEK: "lastWeek",
  NEXT_MONTH: "nextMonth",
  THIS_MONTH: "thisMonth",
  LAST_MONTH: "lastMonth",
  NEXT_QUARTER: "nextQuarter",
  THIS_QUARTER: "thisQuarter",
  LAST_QUARTER: "lastQuarter",
  NEXT_YEAR: "nextYear",
  THIS_YEAR: "thisYear",
  LAST_YEAR: "lastYear",
  YEAR_TO_DATE: "yearToDate",
  Q1: "Q1",
  Q2: "Q2",
  Q3: "Q3",
  Q4: "Q4",
  M1: "M1",
  M2: "M2",
  M3: "M3",
  M4: "M4",
  M5: "M5",
  M6: "M6",
  M7: "M7",
  M8: "M8",
  M9: "M9",
  M10: "M10",
  M11: "M11",
  M12: "M12",
} as const;

export type PivotFilterType = (typeof PivotFilterType)[keyof typeof PivotFilterType];

/** A single pivot filter (CT_PivotFilter) */
export interface PivotFilterOptions {
  /** Field index to filter on (required) */
  readonly fld: number;
  /** Filter type (required) */
  readonly type: PivotFilterType;
  /** Filter ID — unique within this pivot table (required) */
  readonly id: number;
  /** Measure field index for OLAP filters */
  readonly mpFld?: number;
  /** Evaluation order */
  readonly evalOrder?: number;
  /** Measure hierarchy */
  readonly iMeasureHier?: number;
  /** Measure field */
  readonly iMeasureFld?: number;
  /** Filter name */
  readonly name?: string;
  /** Filter description */
  readonly description?: string;
  /** First string value for caption/date filters */
  readonly stringValue1?: string;
  /** Second string value for between filters */
  readonly stringValue2?: string;
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
  /** Pivot filters (CT_PivotFilters) */
  readonly filters?: readonly PivotFilterOptions[];
  /** Field names to use as page/report filters */
  readonly pages?: readonly string[];
}

/** OLAP properties for pivot cache (CT_OlapPr) */
export interface OlapPrOptions {
  /** Local cube connection string */
  readonly local?: string;
  /** Local connection string */
  readonly localConnection?: string;
  /** Send locale info to OLAP server */
  readonly sendLocale?: boolean;
  /** Row dimensions */
  readonly rowDrillCount?: number;
  /** Column dimensions */
  readonly colDrillCount?: number;
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
