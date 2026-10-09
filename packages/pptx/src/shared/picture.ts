import type {
  ShapeType,
  BasePictureOptions,
  BlipCompression,
  BlackWhiteMode,
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
 * Picture (p:pic) for pptx slides. Extends cross-format BasePictureOptions
 * (binary data + cNvPr name/description/title/hidden → p:cNvPr) with absolute
 * EMU positioning and optional shape-level effects.
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
  /** Rotate image fill with the shape frame (p:blipFill @rotWithShape). */
  rotWithShape?: boolean;
  type: "png" | "jpg" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf";
  /**
   * Media file name inside the package (ppt/media/<fileName>). Round-trip
   * keeps the source name; fresh generation derives it from the shape name.
   */
  fileName?: string;
  /**
   * Source `a:blip` `@r:embed` when the target has no image bytes
   * (round-trip only — a broken internal reference).
   */
  relationshipId?: string;
  /** Compression state (a:blip @cstate); absent = attribute omitted. */
  compression?: BlipCompression;
  /** Shape-level effects on p:spPr (e.g. shadow/reflection). */
  effects?: EffectListOptions;
  /** Raw a:extLst inner XML on p:spPr; round-trip only. */
  shapePropertiesExt?: string;
  /** Crop rectangle (a:srcRect) — integer percent insets. */
  sourceRectangle?: SourceRectangleOptions;
  /**
   * Fill rectangle inside the blip fill's a:stretch: undefined emits the
   * standard empty `<a:fillRect/>`, false the bare `<a:stretch/>` some
   * writers produce, an object carries inset percents.
   */
  fillRectangle?: SourceRectangleOptions | false;
  /** Blip color effects (a:lum, a:duotone, … children of a:blip). */
  blipEffects?: BlipEffectsOptions;
  /** Local-DPI display hint (a14:useLocalDpi in the a:blip extension list). */
  useLocalDpi?: boolean;
  /** Verbatim a:blip extension-list children beyond modeled hints (round-trip only). */
  blipExt?: string;
  /** Fill on p:spPr (a:noFill on cropped pictures is common). */
  fill?: FillOptions;
  /** Outline on p:spPr (a:ln — decorated pictures carry one). */
  outline?: OutlineOptions;
  /**
   * Preset geometry on p:spPr. Fresh pictures always carry a rect frame;
   * null suppresses the element for sources that omit it.
   */
  geometry?: ShapeType | null;
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
  /** @bwMode on p:spPr — black-and-white render mode. */
  blackWhiteMode?: BlackWhiteMode;
}
