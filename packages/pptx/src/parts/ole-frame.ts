/**
 * OLE Object frame types for PresentationML.
 *
 * @module
 */

import type {
  BlipCompression,
  DataType,
  GraphicFrameLockingOptions,
  NonVisualDrawingPropertiesOptions,
  UniversalMeasure,
} from "@office-open/core";
import type { SourceRectangleOptions } from "@office-open/core/drawing";
import type { FillOptions, OutlineOptions } from "@office-open/core/drawing";
import type { ShapeType } from "@office-open/core/drawing";
import type { GraphicFrameNvPrOptions } from "@parts/descriptors/graphic-frame";

// ── Options ──

export interface OleEmbedOptions {
  /** OLE container binary — registered as ppt/embeddings/oleObjectN.bin. */
  data: DataType;
  /** Source embedding file name (ppt/embeddings/*) — round-trip only. */
  fileName?: string;
  /** Source embedding relationship kind; undefined (fresh authoring) = oleObject. */
  relationshipType?: "oleObject" | "package";
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
  /** Embedded picture cNvPr `@id`; source identity is preserved on round-trip. */
  id?: number;
  /** Embedded picture cNvPr `@name`; source identity is preserved on round-trip. */
  name?: string;
  /** Source icon media file name, pinned for round-trip relationship wiring. */
  fileName?: string;
  /** false preserves a source icon blip fill with no `a:stretch` child. */
  stretchPresent?: boolean;
  /** Source `a:fillRect`; `false` preserves bare `<a:stretch/>`. */
  fillRectangle?: SourceRectangleOptions | false;
  /** Embedded picture fill (`a:blipFill/p:spPr`); absence inherits the theme. */
  fill?: FillOptions;
  /** Embedded picture outline (`a:ln`). */
  outline?: OutlineOptions;
  /** false preserves a source embedded picture with no `p:spPr`. */
  shapePropertiesPresent?: boolean;
  /** Embedded picture geometry; `null` preserves an omitted `a:prstGeom`. */
  geometry?: ShapeType | null;
  /** false preserves a source embedded picture with no `a:xfrm`. */
  transformPresent?: boolean;
  /** Embedded picture transform in EMU. */
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  /** Compression state (a:blip `@cstate`); absent = attribute omitted. */
  compression?: BlipCompression;
}

/**
 * OLE object frame (p:graphicFrame with p:oleObj) for pptx slides. cNvPr
 * fields (name/description/title/hidden) from
 * NonVisualDrawingPropertiesOptions.
 */
export interface OleOptions extends NonVisualDrawingPropertiesOptions, GraphicFrameNvPrOptions {
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
  /**
   * p:oleObj `@name` — independent of the frame's p:cNvPr `@name`. `false`
   * preserves a source object that omitted this optional attribute; fresh
   * authoring may omit the field to use the frame-name default.
   */
  objectName?: string | false;
  shapeId?: string;
  showAsIcon?: boolean;
  /** Image width (EMU) for icon/preview (AG_Ole `@imgW`) */
  imageWidth?: number;
  /** Image height (EMU) for icon/preview (AG_Ole `@imgH`) */
  imageHeight?: number;
  /**
   * Embedded OLE object (binary registered as ppt/embeddings/oleObjectN.bin).
   * Mutually exclusive with `link` — p:oleObj's content model is a
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
