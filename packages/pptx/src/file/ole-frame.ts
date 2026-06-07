/**
 * OLE Object frame types for PresentationML.
 *
 * @module
 */

// ── Options ──

export interface OleEmbedOptions {
  /** Relationship ID for the embedded OLE data */
  readonly rId: string;
  /** Last update in document (ISO 8601) */
  readonly lastEdited?: string;
}

export interface OleLinkOptions {
  /** Relationship ID for the linked OLE data */
  readonly rId: string;
  /** Last update in document (ISO 8601) */
  readonly updateLastEdited?: string;
  /** Automatic or manual update */
  readonly autoUpdate?: boolean;
}

export interface OleOptions {
  /** Position and size */
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  /** OLE program ID (e.g., "Excel.Sheet.12") */
  readonly progId?: string;
  /** Shape ID */
  readonly spid?: string;
  /** Object name */
  readonly name?: string;
  /** Show as icon */
  readonly showAsIcon?: boolean;
  /** Image width (EMU) for icon/preview */
  readonly imgW?: number;
  /** Image height (EMU) for icon/preview */
  readonly imgH?: number;
  /** Embed mode (provides rId for embedded OLE) */
  readonly embed?: OleEmbedOptions;
  /** Link mode (provides rId for linked OLE) */
  readonly link?: OleLinkOptions;
  /** Relationship ID for the preview/icon image */
  readonly imgRId?: string;
  /** Follow color scheme: "none", "full", or "textAndBackground" */
  readonly followColorScheme?: "none" | "full" | "textAndBackground";
}

export interface OleData {
  readonly key: string;
  readonly rId: string;
  readonly progId?: string;
}
