/**
 * Embedded Excel workbook generator for chart data using @office-open/xlsx.
 *
 * Generates xlsx files for embedding chart data in Word documents.
 *
 * @module
 */

import type { ChartSeriesData, BubbleSeriesData } from "./types";

/** Data structure for embedded workbook */
export interface EmbeddedWorkbookData {
  categories: readonly string[];
  series: readonly ChartSeriesData[] | readonly BubbleSeriesData[];
}

/**
 * Generate an embedded Excel workbook (xlsx) for chart data.
 *
 * Creates an xlsx file with a single worksheet containing the chart's
 * categories and series data arranged in a table format suitable for
 * Word's chart editing feature.
 *
 * Note: This function returns a Promise that resolves to the workbook options
 * structure. The actual xlsx generation should be done by the caller using
 * @office-open/xlsx's generateWorkbookSync.
 *
 * @param data - Chart data (categories and series)
 * @returns WorkbookOptions for xlsx generation
 */
export function buildEmbeddedWorkbookOptions(data: EmbeddedWorkbookData): {
  worksheets: Array<{
    name: string;
    rows: Array<{
      cells: Array<{ value: string | number | null }>;
    }>;
  }>;
} {
  const { categories, series } = data;

  // Build rows
  const rows: Array<{ cells: Array<{ value: string | number | null }> }> = [];

  // Row 1: Header row - empty cell + category names
  const headerCells: Array<{ value: string | number | null }> = [
    { value: "" }, // Empty top-left cell
  ];
  for (const cat of categories) {
    headerCells.push({ value: cat });
  }
  rows.push({ cells: headerCells });

  // Data rows: Series name + values
  for (const s of series as ChartSeriesData[]) {
    const cells: Array<{ value: string | number | null }> = [
      { value: s.name }, // Series name in first column
    ];
    for (const val of s.values) {
      cells.push({ value: val });
    }
    rows.push({ cells });
  }

  return {
    worksheets: [
      {
        name: "Sheet1",
        rows,
      },
    ],
  };
}
