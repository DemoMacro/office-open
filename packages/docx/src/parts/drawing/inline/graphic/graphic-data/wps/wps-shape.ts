import type {
  CustomGeometryOptions,
  EffectDagOptions,
  EffectListOptions,
  FillOptions,
  OutlineOptions,
  PresetGeometryOptions,
  Scene3DOptions,
  Shape3DOptions,
  SolidFillOptions,
} from "@office-open/core/drawingml";
import type { ParagraphOptions } from "@parts/paragraph/paragraph";

/**
 * Shape types for WordprocessingML drawings.
 *
 * Provides type definitions for shape elements in DrawingML.
 *
 * @module
 */
import type { BodyPropertiesOptions } from "./body-properties";
import type { NonVisualShapePropertiesOptions } from "./non-visual-shape-properties";

/**
 * A style-matrix reference (CT_StyleMatrixReference): an `idx` into the theme
 * style matrix plus an optional color override.
 */
export interface StyleMatrixReferenceOptions {
  /** Style index ("0","1","2",…; or "minor"/"major" for font references). */
  idx: string;
  /** Color (EG_ColorChoice) — schemeClr/srgbClr/hslClr/sysClr/prstClr/scRgbClr. */
  color?: SolidFillOptions;
}

/**
 * Shape style (CT_ShapeStyle): line/fill/effect/font references into the
 * document's theme. Word emits this for every shape that inherits theme styling.
 */
export interface ShapeStyleOptions {
  lineReference?: StyleMatrixReferenceOptions;
  fillReference?: StyleMatrixReferenceOptions;
  effectReference?: StyleMatrixReferenceOptions;
  fontReference?: StyleMatrixReferenceOptions;
}

export interface ShapeCoreOptions {
  children: (ParagraphOptions | string)[];
  nonVisualProperties?: NonVisualShapePropertiesOptions;
  bodyProperties?: BodyPropertiesOptions;
  outline?: OutlineOptions;
  fill?: FillOptions;
  customGeometry?: CustomGeometryOptions;
  presetGeometry?: PresetGeometryOptions;
  effectDag?: EffectDagOptions;
  effects?: EffectListOptions;
  scene3d?: Scene3DOptions;
  shape3d?: Shape3DOptions;
  /** Theme style references (wps:style → lnRef/fillRef/effectRef/fontRef). */
  style?: ShapeStyleOptions;
  /**
   * Linked text box chain (wps:linkedTxbx) — the shape's text lives in the
   * linked part instead of inline w:txbxContent. XSD choice: exclusive with
   * inline text content.
   */
  linkedTextBox?: {
    /** Chain id shared by all boxes in the link (@id, required). */
    id: number;
    /** Position of this box in the chain (@seq, required). */
    sequence: number;
  };
  /** East-Asian vertical text flow (wps:wsp @normalEastAsianFlow, default false). */
  normalEastAsianFlow?: boolean;
}
