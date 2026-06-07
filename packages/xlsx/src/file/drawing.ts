/**
 * XLSX Drawing types — image and chart anchor options for worksheets.
 *
 * @module
 */

export interface ImageOptions {
  /** 1-based column */
  col: number;
  /** Column offset in EMU (default 0) */
  colOffset?: number;
  /** 1-based row */
  row: number;
  /** Row offset in EMU (default 0) */
  rowOffset?: number;
  /** Relationship ID for the image */
  rId: string;
  /** Lock anchor with sheet (default true) */
  locksWithSheet?: boolean;
  /** Print with sheet (default true) */
  printsWithSheet?: boolean;
}

export interface ChartAnchorOptions {
  /** 1-based column */
  col: number;
  /** Column offset in EMU (default 0) */
  colOffset?: number;
  /** 1-based row */
  row: number;
  /** Row offset in EMU (default 0) */
  rowOffset?: number;
  /** Relationship ID for the chart */
  rId: string;
  /** Lock anchor with sheet (default true) */
  locksWithSheet?: boolean;
  /** Print with sheet (default true) */
  printsWithSheet?: boolean;
}
