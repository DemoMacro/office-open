import type {
  CustomGeometryOptions,
  EffectDagOptions,
  EffectListOptions,
  FillOptions,
  OutlineOptions,
  PresetGeometryOptions,
  ShapePropertiesExtensionOptions,
  Scene3DOptions,
  Shape3DOptions,
  ShapeType,
  SolidFillOptions,
} from "@office-open/core/drawing";
import type { ParagraphOptions } from "@parts/paragraph/paragraph";
import type { SectionChild } from "@shared/section";

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
 * A style-matrix reference (CT_StyleMatrixReference): an index into the theme
 * style matrix plus an optional color override.
 */
export interface ShapeStyleReferenceOptions {
  /** Index into the theme style matrix list (idx attribute). */
  index: number;
  /** Color (EG_ColorChoice) — schemeClr/srgbClr/hslClr/sysClr/prstClr/scRgbClr. */
  color?: SolidFillOptions;
}

/**
 * Shape style (CT_ShapeStyle): line/fill/effect/font references into the
 * document's theme. Word emits this for every shape that inherits theme styling.
 */
export interface ShapeStyleOptions {
  lineReference?: ShapeStyleReferenceOptions;
  fillReference?: ShapeStyleReferenceOptions;
  effectReference?: ShapeStyleReferenceOptions;
  /** a:fontRef — @idx is ST_FontCollectionIndex, not a number. */
  fontReference?: { collection: "major" | "minor" | "none"; color?: SolidFillOptions };
}

export type ShapeTextBoxChild = ParagraphOptions | string | SectionChild;

/** A Word 2010 text-box content part referenced by `wps:txbx/@r:txbx`. */
export interface TextBoxPartOptions {
  /** Package path of the text-box part (e.g. `word/txbx1.xml`). */
  path: string;
  /** Text-box sequence (`wps:txbx/@txbxSeq`). */
  sequence: number;
}

export interface ShapeCoreOptions {
  /**
   * Block-level w:txbxContent children (w:EG_BlockLevelElts). Paragraphs keep
   * their established shorthand; wrapped paragraphs and other block elements
   * use SectionChild variants.
   */
  children: ShapeTextBoxChild[];
  nonVisualProperties?: NonVisualShapePropertiesOptions;
  bodyProperties?: BodyPropertiesOptions;
  outline?: OutlineOptions;
  fill?: FillOptions;
  customGeometry?: CustomGeometryOptions;
  /** Preset geometry (a:prstGeom) — a bare string is shorthand for `{ preset: "<name>" }`. */
  geometry?: ShapeType | PresetGeometryOptions;
  effectDag?: EffectDagOptions;
  effects?: EffectListOptions;
  scene3d?: Scene3DOptions;
  shape3d?: Shape3DOptions;
  /** Shape-property extensions (wps:spPr/a:extLst/a:ext). */
  extensions?: ShapePropertiesExtensionOptions[];
  /** Theme style references (wps:style → lnRef/fillRef/effectRef/fontRef). */
  style?: ShapeStyleOptions;
  /**
   * External Word 2010 text-box content part. Round-trips `wps:txbx/@r:txbx`
   * while the part bytes remain in DocumentOptions.rawParts.
   */
  textBoxPart?: TextBoxPartOptions;
  /**
   * Linked text box chain (wps:linkedTxbx) — the shape's text lives in the
   * linked part instead of inline w:txbxContent. XSD choice: exclusive with
   * inline text content.
   */
  linkedTextBox?: {
    /** Chain id shared by all boxes in the link (`@id`, required). */
    id: number;
    /** Position of this box in the chain (`@seq`, required). */
    sequence: number;
  };
  /** East-Asian vertical text flow (wps:wsp `@normalEastAsianFlow`, default false). */
  normalEastAsianFlow?: boolean;
}
