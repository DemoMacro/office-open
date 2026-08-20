import type {
  NonVisualDrawingPropertiesOptions,
  ShapeLockingOptions,
  UniversalMeasure,
} from "@office-open/core";
import type { ReadContext } from "@office-open/core/descriptor";
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
import { parseColorChoice, isPlainRgbColor } from "@office-open/core/drawing";
import type { SolidFillOptions } from "@office-open/core/drawing";
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

/** Style-matrix reference color: hex shorthand or a full color choice
 * (schemeClr with transforms, sysClr, hslClr, …). */
/** CT_Placeholder `@type` — ST_PlaceholderType. */
export type PlaceholderType =
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

/** CT_Placeholder `@sz` — sizing hint. */
export type PlaceholderSize = "full" | "half" | "quarter";

/** CT_Placeholder `@orient` — orientation hint. */
export type PlaceholderOrientation = "horz" | "vert";

export type StyleReferenceColor = string | SolidFillOptions;

export interface ShapeStyleOptions {
  lineReference?: { index: number; color?: StyleReferenceColor };
  fillReference?: { index: number; color?: StyleReferenceColor };
  effectReference?: { index: number; color?: StyleReferenceColor };
  /** a:fontRef — @idx is ST_FontCollectionIndex (major/minor/none), not a number. */
  fontReference?: { index: number | "major" | "minor" | "none"; color?: StyleReferenceColor };
}

export interface ShapeOptions extends NonVisualDrawingPropertiesOptions {
  id?: number;
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  geometry?: string | PresetGeometryOptions;
  customGeometry?: CustomGeometryOptions;
  /**
   * `null` marks a source spPr with no fill child — absence is the fidelity
   * (the shape inherits its fill), so stringify emits nothing instead of the
   * fresh-authoring noFill default.
   */
  fill?: FillOptions | null;
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
  placeholder?: PlaceholderType;
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
 * so refs without it are skipped. Each ref carries a bare EG_ColorChoice
 * child (any of the six color kinds).
 */
export function readShapeStyle(styleEl: XmlElement, ctx: ReadContext): ShapeStyleOptions {
  const style: ShapeStyleOptions = {};
  const lnRef = findChild(styleEl, "a:lnRef");
  if (lnRef) {
    const idx = attrNum(lnRef, "idx");
    if (idx !== undefined) {
      const color = readRefColor(lnRef, ctx);
      style.lineReference = color !== undefined ? { index: idx, color } : { index: idx };
    }
  }
  const fillRef = findChild(styleEl, "a:fillRef");
  if (fillRef) {
    const idx = attrNum(fillRef, "idx");
    if (idx !== undefined) {
      const color = readRefColor(fillRef, ctx);
      style.fillReference = color !== undefined ? { index: idx, color } : { index: idx };
    }
  }
  const effectRef = findChild(styleEl, "a:effectRef");
  if (effectRef) {
    const idx = attrNum(effectRef, "idx");
    if (idx !== undefined) {
      const color = readRefColor(effectRef, ctx);
      style.effectReference = color !== undefined ? { index: idx, color } : { index: idx };
    }
  }
  const fontRef = findChild(styleEl, "a:fontRef");
  if (fontRef) {
    // @idx is ST_FontCollectionIndex — numeric stays numeric, the
    // major/minor/none tokens keep their string form.
    const raw = attr(fontRef, "idx");
    let idx: number | "major" | "minor" | "none" | undefined;
    if (raw === "major" || raw === "minor" || raw === "none") idx = raw;
    else idx = attrNum(fontRef, "idx");
    if (idx !== undefined) {
      const color = readRefColor(fontRef, ctx);
      style.fontReference = color !== undefined ? { index: idx, color } : { index: idx };
    }
  }
  return style;
}

/** Read a ref's EG_ColorChoice; a bare srgbClr collapses to the hex shorthand,
 * scheme/sys/hsl colors keep their structured form. */
function readRefColor(refEl: XmlElement, ctx: ReadContext): StyleReferenceColor | undefined {
  const color = parseColorChoice(refEl, ctx);
  if (!color || Object.keys(color).length === 0) return undefined;
  if (isPlainRgbColor(color)) return (color as { value: string }).value;
  return color;
}
