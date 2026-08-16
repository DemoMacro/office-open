import type { BasePictureOptions, EffectListOptions, UniversalMeasure } from "@office-open/core";
import type { TextHyperlinkOptions } from "@office-open/core/drawing";

/**
 * Picture (p:pic) options for PPTX slides.
 *
 * Extends the cross-format {@link BasePictureOptions} (binary data + non-visual
 * drawing properties) with absolute EMU positioning and optional shape-level
 * effects. The base cNvPr fields (name/description/title/hidden) flow straight
 * through to p:cNvPr. The single source of truth for both the public
 * slide-child entry and the descriptor.
 */
export interface PictureOptions extends BasePictureOptions {
  /** Picture id (p:cNvPr `@id`). Auto-generated if omitted. */
  id?: number;
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  /** Flip horizontally (a:xfrm `@flipH`). */
  flipHorizontal?: boolean;
  /** Flip vertically (a:xfrm `@flipV`). */
  flipVertical?: boolean;
  /** Rotation angle in degrees (e.g., 45 = 45°). */
  rotation?: number;
  type: "png" | "jpg" | "gif" | "bmp" | "emf" | "wmf";
  /** Shape-level effects on p:spPr (e.g. shadow/reflection). */
  effects?: EffectListOptions;
  /**
   * Click hyperlink on the picture itself (a:hlinkClick inside p:cNvPr) —
   * jump to a URL or another slide when the picture is clicked.
   */
  hyperlink?: TextHyperlinkOptions;
}
