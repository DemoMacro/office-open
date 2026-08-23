import type { PositiveUniversalMeasure } from "@office-open/core";

/**
 * Page orientation (ST_PageOrientation): "portrait" (default) or "landscape"
 * — landscape swaps the effective paper width/height when printing.
 *
 * @publicApi
 */
export const PageOrientation = {
  PORTRAIT: "portrait",
  LANDSCAPE: "landscape",
} as const;

export interface PageSizeProperties {
  /** Page width in twips (15840 = 11"). */
  width?: number | PositiveUniversalMeasure;
  /** Page height in twips (12240 = 8.5"). */
  height?: number | PositiveUniversalMeasure;
  /** Page orientation; when landscape, width/height are printed swapped on the paper. Default portrait. */
  orientation?: (typeof PageOrientation)[keyof typeof PageOrientation];
  /** Printer-specific paper code, stored verbatim (printer picks the tray/paper type). */
  code?: number;
}
