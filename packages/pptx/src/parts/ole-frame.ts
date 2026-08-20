/**
 * OLE Object frame types for PresentationML.
 *
 * @module
 */

import type {
  DataType,
  GraphicFrameLockingOptions,
  NonVisualDrawingPropertiesOptions,
  UniversalMeasure,
} from "@office-open/core";
import type { NvPrPlaceholderOptions } from "@parts/descriptors/graphic-frame";

// ── Options ──

export interface OleEmbedOptions {
  /** OLE container binary — registered as ppt/embeddings/oleObjectN.bin. */
  data: DataType;
  /** Follow color scheme (p:embed `@followColorScheme`) */
  followColorScheme?: "none" | "full" | "textAndBackground";
}

export interface OleLinkOptions {
  /**
   * Linked OLE source URL — registered as an External oleObject relationship
   * of the owning slide/layout when the object is emitted.
   */
  url: string;
  /** Automatic or manual update (p:link `@updateAutomatic`) */
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
export interface OleOptions extends NonVisualDrawingPropertiesOptions, NvPrPlaceholderOptions {
  /** Frame locking (a:graphicFrameLocks). undefined = fresh default
   * (noGrp="1"); null = empty cNvGraphicFramePr; object = explicit flags. */
  locking?: GraphicFrameLockingOptions | null;
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
  /** Image width (EMU) for icon/preview (AG_Ole `@imgW`) */
  imageWidth?: number;
  /** Image height (EMU) for icon/preview (AG_Ole `@imgH`) */
  imageHeight?: number;
  /**
   * Embedded OLE object (binary registered as ppt/embeddings/oleObjectN.bin).
   * Mutually exclusive with {@link link} — p:oleObj's content model is a
   * required choice between p:embed and p:link.
   */
  embed?: OleEmbedOptions;
  /**
   * Linked OLE object (external source URL, no bytes in the package);
   * mutually exclusive with embed.
   */
  link?: OleLinkOptions;
  /**
   * Icon/preview image (p:pic under p:oleObj). MS Office refuses to open a
   * presentation whose oleObj has no picture, so this is effectively required
   * for embedded objects.
   */
  iconImage?: OleIconImageOptions;
}
