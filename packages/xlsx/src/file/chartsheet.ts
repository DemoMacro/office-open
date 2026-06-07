/**
 * Chartsheet options for SpreadsheetML documents.
 *
 * A chartsheet is a worksheet that contains only a chart (no cells).
 *
 * Reference: OOXML transitional, sml.xsd, CT_Chartsheet
 *
 * @module
 */

export interface ChartsheetPageMargins {
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
  header?: number;
  footer?: number;
}

export interface ChartsheetPageSetup {
  /** Paper size (1=Letter, 9=A4, etc.) */
  paperSize?: number;
  /** Orientation ("default" | "portrait" | "landscape") */
  orientation?: string;
  /** Horizontal DPI */
  horizontalDpi?: number;
  /** Vertical DPI */
  verticalDpi?: number;
  /** Copies to print */
  copies?: number;
}

export interface ChartsheetProtectionOptions {
  /** Content is protected */
  content?: boolean;
  /** Objects are protected */
  objects?: boolean;
}

export interface ChartsheetHeaderFooterOptions {
  /** Different first page header/footer */
  differentFirst?: boolean;
  /** Different odd/even page headers/footers */
  differentOddEven?: boolean;
  /** Odd page header */
  oddHeader?: string;
  /** Odd page footer */
  oddFooter?: string;
}

export interface ChartsheetOptions {
  /** Sheet name */
  name?: string;
  /** Tab color (hex ARGB, e.g. "FF4472C4") */
  tabColor?: string;
  /** Page margins */
  pageMargins?: ChartsheetPageMargins;
  /** Page setup */
  pageSetup?: ChartsheetPageSetup;
  /** Header/footer */
  headerFooter?: ChartsheetHeaderFooterOptions;
  /** Sheet protection */
  sheetProtection?: ChartsheetProtectionOptions;
  /** Published to server (CT_ChartsheetPr @published) */
  published?: boolean;
  /** Zoom to fit (CT_ChartsheetView @zoomToFit) */
  zoomToFit?: boolean;
  /** Chart definition (type, title, series, etc.) */
  chart: {
    type: string;
    title?: string;
    categories?: string[];
    series: {
      name: string;
      values: number[];
    }[];
  };
}
