/**
 * Frameset and Frame type definitions for webSettings.xml.
 *
 * Framesets define a frames-based layout within a Word document (web layout).
 * XML generation is handled by the descriptor at compile/descriptors/web-settings.ts.
 *
 * Reference: OOXML transitional, wml.xsd, CT_Frameset / CT_Frame
 *
 * @module
 */

// ── Options ──

export interface FrameOptions {
  /** Frame size (e.g., "50%") */
  size?: string;
  /** Frame name */
  name?: string;
  /** Frame title */
  title?: string;
  /** Source file link rId */
  sourceRId?: string;
  /** Margin width in pixels */
  marginWidth?: number;
  /** Margin height in pixels */
  marginHeight?: number;
  /** Scrollbar mode */
  scrollbar?: "on" | "off" | "auto";
  /** No resize allowed */
  noResizeAllowed?: boolean;
  /** Linked to file */
  linkedToFile?: boolean;
  /** Long description relationship ID */
  longDescRId?: string;
}

export interface FramesetSplitbarOptions {
  /** Splitbar width in twips */
  width?: number;
  /** Splitbar color (hex, e.g., "auto") */
  color?: string;
  /** No border */
  noBorder?: boolean;
  /** Flat borders */
  flatBorders?: boolean;
}

export interface FramesetOptions {
  /** Frameset size (e.g., "100%") */
  size?: string;
  /** Splitbar configuration */
  splitbar?: FramesetSplitbarOptions;
  /** Layout direction */
  layout?: "rows" | "cols";
  /** Frameset title */
  title?: string;
  /** Child framesets and frames (in order) */
  children?: (FramesetOptions | FrameOptions)[];
}
