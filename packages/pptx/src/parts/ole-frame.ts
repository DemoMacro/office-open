/**
 * OLE Object frame types for PresentationML.
 *
 * @module
 */

import type {
  DataType,
  NonVisualDrawingPropertiesOptions,
  UniversalMeasure,
} from "@office-open/core";

// ── Options ──

export interface OleEmbedOptions {
  /** OLE container binary — registered as ppt/embeddings/oleObjectN.bin. */
  data: DataType;
}

export interface OleLinkOptions {
  /** Relationship ID for the linked OLE data */
  rId: string;
  /** Automatic or manual update */
  autoUpdate?: boolean;
}

export interface OleIconImageOptions {
  /** Icon/preview image bytes (binary or base64 data URL). */
  data: DataType;
  /** Image type / extension (e.g. "png", "emf"). */
  type: string;
}

/**
 * OLE Object frame options for pptx slides (p:graphicFrame with p:oleObj). The
 * cNvPr fields (name/description/title/hidden) come from
 * {@link NonVisualDrawingPropertiesOptions}. The single source of truth for
 * both the public slide-child entry and the descriptor.
 */
export interface OleOptions extends NonVisualDrawingPropertiesOptions {
  /** OLE frame id (p:cNvPr `@id`). Auto-generated if omitted. */
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
  /** Embedded OLE object (binary registered as ppt/embeddings/oleObjectN.bin). */
  embed?: OleEmbedOptions;
  /** Link mode (provides rId for linked OLE data) */
  link?: OleLinkOptions;
  /**
   * Icon/preview image (p:pic under p:oleObj). MS Office refuses to open a
   * presentation whose oleObj has no picture, so this is effectively required
   * for embedded objects.
   */
  iconImage?: OleIconImageOptions;
  /** Follow color scheme (p:embed/`@followColorScheme`): "none", "full", or "textAndBackground" */
  followColorScheme?: "none" | "full" | "textAndBackground";
}
