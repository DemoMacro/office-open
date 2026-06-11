/**
 * OLE Object frame types for PresentationML.
 *
 * @module
 */

import type { UniversalMeasure } from "@office-open/core";

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

export interface OleOptions {
  /** Position and size */
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  /** OLE program ID (e.g., "Excel.Sheet.12") */
  progId?: string;
  /** Shape ID */
  spid?: string;
  /** Object name */
  name?: string;
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
