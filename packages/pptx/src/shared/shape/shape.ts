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
} from "@office-open/core/drawing";

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
  /** Rotation angle in degrees (e.g., 45 = 45°). */
  rotation?: number;
  textBody?: TextBodyOptions;
  locking?: ShapeLockingOptions;
  /** CT_Placeholder @type — ST_PlaceholderType. */
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
  /** CT_Placeholder @sz — sizing hint (default "full"). */
  placeholderSize?: "full" | "half" | "quarter";
  /** CT_Placeholder @orient — orientation hint (default "horz"). */
  placeholderOrientation?: "horz" | "vert";
  useBackgroundFill?: boolean;
  isPhoto?: boolean;
  userDrawn?: boolean;
  hasCustomPrompt?: boolean;
  style?: ShapeStyleOptions;
  blackWhiteMode?:
    | "clr"
    | "auto"
    | "gray"
    | "ltGray"
    | "invGray"
    | "grayWhite"
    | "blackGray"
    | "blackWhite"
    | "black"
    | "white"
    | "hidden";
}
