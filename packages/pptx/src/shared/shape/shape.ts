import type {
  NonVisualDrawingPropertiesOptions,
  ShapeLockingOptions,
  UniversalMeasure,
} from "@office-open/core";
import type {
  PresetGeometryOptions,
  CustomGeometryOptions,
  TextBodyOptions,
  OutlineOptions,
  EffectListOptions,
  Scene3DOptions,
  Shape3DOptions,
  FillOptions,
  BlackWhiteMode,
  TextHyperlinkOptions,
} from "@office-open/core/drawing";
import { attr, attrNum, findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

/**
 * Shape options type for PPTX.
 *
 * The single source of truth for both the public slide-child entry and the
 * descriptor. Shape-level animation is intentionally absent — CT_Shape has no
 * timing child in the XSD; animation lives at the slide level
 * (slide.animations / timing).
 *
 * @module
 */
export interface ShapeStyleOptions {
  lineReference?: { index: number; color?: string };
  fillReference?: { index: number; color?: string };
  effectReference?: { index: number; color?: string };
  fontReference?: { index: number; color?: string };
}

export interface ShapeOptions extends NonVisualDrawingPropertiesOptions {
  id?: number;
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  geometry?: string | PresetGeometryOptions;
  customGeometry?: CustomGeometryOptions;
  fill?: FillOptions;
  outline?: OutlineOptions;
  effects?: EffectListOptions;
  scene3d?: Scene3DOptions;
  shape3d?: Shape3DOptions;
  /** Raw a:extLst inner XML — verbatim round-trip for unmodeled extensions. */
  ext?: string;
  flipHorizontal?: boolean;
  /** Flip vertically (a:xfrm `@flipV`). */
  flipVertical?: boolean;
  /** Rotation angle in degrees (e.g., 45 = 45°). */
  rotation?: number;
  textBody?: TextBodyOptions;
  locking?: ShapeLockingOptions;
  /** CT_Placeholder `@type` — ST_PlaceholderType. */
  placeholder?:
    | "title"
    | "body"
    | "ctrTitle"
    | "subTitle"
    | "dt"
    | "sldNum"
    | "ftr"
    | "hdr"
    | "obj"
    | "chart"
    | "tbl"
    | "clipArt"
    | "dgm"
    | "media"
    | "sldImg"
    | "pic";
  placeholderIndex?: number;
  /** CT_Placeholder `@sz` — sizing hint (default "full"). */
  placeholderSize?: "full" | "half" | "quarter";
  /** CT_Placeholder `@orient` — orientation hint (default "horz"). */
  placeholderOrientation?: "horz" | "vert";
  useBackgroundFill?: boolean;
  isPhoto?: boolean;
  userDrawn?: boolean;
  hasCustomPrompt?: boolean;
  style?: ShapeStyleOptions;
  /**
   * `@bwMode` (ST_BlackWhiteMode) on `p:spPr` — how the shape renders in
   * black-and-white view/print.
   */
  blackWhiteMode?: BlackWhiteMode;
  /**
   * Click hyperlink on the shape itself (a:hlinkClick inside p:cNvPr) —
   * jump to a URL or another slide when the shape is clicked.
   */
  hyperlink?: TextHyperlinkOptions;
}

/**
 * Parse a p:style element (a:lnRef/a:fillRef/a:effectRef/a:fontRef) into
 * ShapeStyleOptions. Shared by the shape descriptor and placeholder
 * inheritance — idx is required per CT_StyleMatrixReference/CT_FontReference,
 * so refs without it are skipped.
 */
export function readShapeStyle(styleEl: XmlElement): ShapeStyleOptions {
  const style: ShapeStyleOptions = {};
  const lnRef = findChild(styleEl, "a:lnRef");
  if (lnRef) {
    const idx = attrNum(lnRef, "idx");
    if (idx !== undefined) {
      const color = readStyleSrgbClr(lnRef);
      style.lineReference = color !== undefined ? { index: idx, color } : { index: idx };
    }
  }
  const fillRef = findChild(styleEl, "a:fillRef");
  if (fillRef) {
    const idx = attrNum(fillRef, "idx");
    if (idx !== undefined) {
      const color = readStyleSrgbClr(fillRef);
      style.fillReference = color !== undefined ? { index: idx, color } : { index: idx };
    }
  }
  const effectRef = findChild(styleEl, "a:effectRef");
  if (effectRef) {
    const idx = attrNum(effectRef, "idx");
    if (idx !== undefined) {
      const color = readStyleSrgbClr(effectRef);
      style.effectReference = color !== undefined ? { index: idx, color } : { index: idx };
    }
  }
  const fontRef = findChild(styleEl, "a:fontRef");
  if (fontRef) {
    const idx = attrNum(fontRef, "idx");
    if (idx !== undefined) {
      const color = readStyleSrgbClr(fontRef);
      style.fontReference = color !== undefined ? { index: idx, color } : { index: idx };
    }
  }
  return style;
}

/** Read an srgbClr val from a style ref (direct child or nested in a:solidFill). */
function readStyleSrgbClr(refEl: XmlElement): string | undefined {
  const direct = findChild(refEl, "a:srgbClr");
  if (direct) return attr(direct, "val");
  const solidFill = findChild(refEl, "a:solidFill");
  if (solidFill) {
    const inner = findChild(solidFill, "a:srgbClr");
    if (inner) return attr(inner, "val");
  }
  return undefined;
}
