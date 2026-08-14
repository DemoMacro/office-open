/**
 * OLE Object frame types for PresentationML.
 *
 * @module
 */

import type { NonVisualDrawingPropertiesOptions, UniversalMeasure } from "@office-open/core";

// ── Options ──

export interface OleEmbedOptions {
  /** Relationship ID for the embedded OLE data */
  rId: string;
  /** Last update in document (ISO 8601) */
  lastEdited?: string;
}

export interface OleLinkOptions {
  /** Relationship ID for the linked OLE data */
  rId: string;
  /** Last update in document (ISO 8601) */
  updateLastEdited?: string;
  /** Automatic or manual update */
  autoUpdate?: boolean;
}

/**
 * OLE Object frame options for pptx slides (p:graphicFrame with p:oleObj). The
 * cNvPr fields (name/description/title/hidden) come from
 * {@link NonVisualDrawingPropertiesOptions}. The single source of truth for
 * both the public slide-child entry and the descriptor.
 */
export interface OleOptions extends NonVisualDrawingPropertiesOptions {
  /** OLE frame id (p:cNvPr @id). Auto-generated if omitted. */
  id?: number;
  /** Position and size */
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  /** OLE program ID (e.g., "Excel.Sheet.12") */
  progId?: string;
  /** Shape ID */
  shapeId?: string;
  /** Show as icon */
  showAsIcon?: boolean;
  /** Image width (EMU) for icon/preview */
  imgW?: number;
  /** Image height (EMU) for icon/preview */
  imgH?: number;
  /** Embed mode (provides rId for embedded OLE) */
  embed?: OleEmbedOptions;
  /** Link mode (provides rId for linked OLE) */
  link?: OleLinkOptions;
  /** Relationship ID for the preview/icon image */
  imgRId?: string;
  /** Follow color scheme: "none", "full", or "textAndBackground" */
  followColorScheme?: "none" | "full" | "textAndBackground";
}

export interface OleData {
  key: string;
  rId: string;
  progId?: string;
}
