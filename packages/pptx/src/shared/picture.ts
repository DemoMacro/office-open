import type {
  BasePictureOptions,
  EffectListOptions,
  PictureLockingOptions,
  UniversalMeasure,
} from "@office-open/core";
import type {
  BlipEffectsOptions,
  FillOptions,
  OutlineOptions,
  Scene3DOptions,
  Shape3DOptions,
  SourceRectangleOptions,
} from "@office-open/core/drawing";
import type { TextHyperlinkOptions } from "@office-open/core/drawing";
import type { NvPrPlaceholderOptions } from "@parts/descriptors/graphic-frame";
import type { ShapeStyleOptions } from "@shared/shape/shape";

/**
 * Picture (p:pic) options for PPTX slides.
 *
 * Extends the cross-format {@link BasePictureOptions} (binary data + non-visual
 * drawing properties) with absolute EMU positioning and optional shape-level
 * effects. The base cNvPr fields (name/description/title/hidden) flow straight
 * through to p:cNvPr. The single source of truth for both the public
 * slide-child entry and the descriptor.
 */
export interface PictureOptions extends BasePictureOptions, NvPrPlaceholderOptions {
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
  /** Crop rectangle (a:srcRect) — integer percent insets. */
  sourceRectangle?: SourceRectangleOptions;
  /** Blip color effects (a:lum, a:duotone, … children of a:blip). */
  blipEffects?: BlipEffectsOptions;
  /** Fill on p:spPr (a:noFill on cropped pictures is common). */
  fill?: FillOptions;
  /** Outline on p:spPr (a:ln — decorated pictures carry one). */
  outline?: OutlineOptions;
  /**
   * Preset geometry on p:spPr. Fresh pictures always carry a rect frame;
   * null suppresses the element for sources that omit it.
   */
  geometry?: "rect" | null;
  /** Picture locks (a:picLocks inside p:cNvPicPr). */
  locking?: PictureLockingOptions;
  /**
   * Click hyperlink on the picture itself (a:hlinkClick inside p:cNvPr) —
   * jump to a URL or another slide when the picture is clicked.
   */
  hyperlink?: TextHyperlinkOptions;
  /** 3D scene (a:scene3d) inside p:spPr. */
  scene3d?: Scene3DOptions;
  /** 3D shape properties (a:sp3d) inside p:spPr. */
  shape3d?: Shape3DOptions;
  /** Shape style matrix reference (p:style). */
  style?: ShapeStyleOptions;
}
