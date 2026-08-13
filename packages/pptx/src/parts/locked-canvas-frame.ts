import type { NonVisualDrawingPropertiesOptions, UniversalMeasure } from "@office-open/core";

export interface LockedCanvasShapeOptions {
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  geometry?: string;
  fill?: string;
  /** Simple text content (a:r/a:t); locked-canvas shapes carry plain text only. */
  textBody?: string;
}

/**
 * Locked canvas frame options for pptx slides (p:graphicFrame with
 * lc:lockedCanvas). The cNvPr fields (name/description/title/hidden) come from
 * {@link NonVisualDrawingPropertiesOptions}. The single source of truth for
 * both the public slide-child entry and the descriptor.
 */
export interface LockedCanvasFrameOptions extends NonVisualDrawingPropertiesOptions {
  /** Locked canvas frame id (p:cNvPr @id). Auto-generated if omitted. */
  id?: number;
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  children?: LockedCanvasShapeOptions[];
}
