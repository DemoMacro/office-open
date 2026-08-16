import type {
  NonVisualDrawingPropertiesOptions,
  ShapePropertiesOptions,
  UniversalMeasure,
} from "@office-open/core";

/**
 * Locked-canvas child shape (a:sp). Shape properties (position/geometry/fill/
 * outline/effects) come from the core CT_ShapeProperties model; text is plain
 * string only — locked-canvas shapes carry no rich text.
 */
export interface LockedCanvasShapeOptions extends ShapePropertiesOptions {
  /** Simple text content (a:r/a:t). */
  textBody?: string;
}

/**
 * Locked canvas frame options for pptx slides (p:graphicFrame with
 * lc:lockedCanvas). The cNvPr fields (name/description/title/hidden) come from
 * {@link NonVisualDrawingPropertiesOptions}. The single source of truth for
 * both the public slide-child entry and the descriptor.
 */
export interface LockedCanvasFrameOptions extends NonVisualDrawingPropertiesOptions {
  /** Locked canvas frame id (p:cNvPr `@id`). Auto-generated if omitted. */
  id?: number;
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  children?: LockedCanvasShapeOptions[];
}
