import type { BasePictureOptions, UniversalMeasure } from "@office-open/core";

/**
 * Picture (p:pic) options for PPTX slides.
 *
 * Extends the cross-format {@link BasePictureOptions} (binary data + non-visual
 * drawing properties) with absolute EMU positioning. The base cNvPr fields
 * (name/description/title/hidden) flow straight through to p:cNvPr.
 */
export interface PictureOptions extends BasePictureOptions {
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  type: "png" | "jpg" | "gif" | "bmp" | "emf" | "wmf";
}
