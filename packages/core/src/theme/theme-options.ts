/**
 * Theme options for OOXML documents — full CT_BaseStyles structure.
 *
 * Mirrors a:theme (CT_BaseStyles + objectDefaults + extraClrSchemeLst) so every
 * part of a theme round-trips: color scheme, font scheme, format (style) matrix,
 * object defaults, and extra color schemes.
 *
 * @module
 */

import type {
  BodyPropertiesOptions,
  EffectListOptions,
  FillOptions,
  OutlineOptions,
  Scene3DOptions,
  Shape3DOptions,
  ShapePropertiesOptions,
  SolidFillOptions,
  StyleMatrixReferenceOptions,
  SystemColorOptions,
  TextListStyleOptions,
} from "../drawing";
import type { Panose } from "../util/values";

export type { StyleMatrixReferenceOptions } from "../drawing";

/** Single font slot — a:latin / a:ea / a:cs / a:sym (CT_TextFont). */
export interface TextFontOptions {
  typeface: string;
  panose?: Panose;
  pitchFamily?: number;
  charset?: number;
}

/** Supplemental font for a script range (CT_SupplementalFont). */
export interface SupplementalFontOptions {
  script: string;
  typeface: string;
}

/** Major or minor font collection (CT_FontCollection). */
export interface FontCollectionOptions {
  latin?: TextFontOptions;
  /** East-asian font (a:ea). */
  eastAsian?: TextFontOptions;
  /** Complex-script font (a:cs). */
  complexScript?: TextFontOptions;
  /** Symbol font (a:sym). */
  symbol?: TextFontOptions;
  supplementalFonts?: SupplementalFontOptions[];
}

/** Font scheme (CT_FontScheme). */
export interface FontSchemeOptions {
  majorFont?: FontCollectionOptions;
  minorFont?: FontCollectionOptions;
  name?: string;
}

/**
 * A theme color slot — hex string (emits a:srgbClr) or a structured system
 * color (emits a:sysClr verbatim).
 */
export type SchemeColorValue = string | SystemColorOptions;

/** Color scheme — 12 theme colors (CT_ColorScheme). */
export interface ColorSchemeOptions {
  dark1?: SchemeColorValue;
  light1?: SchemeColorValue;
  dark2?: SchemeColorValue;
  light2?: SchemeColorValue;
  accent1?: SchemeColorValue;
  accent2?: SchemeColorValue;
  accent3?: SchemeColorValue;
  accent4?: SchemeColorValue;
  accent5?: SchemeColorValue;
  accent6?: SchemeColorValue;
  hyperlink?: SchemeColorValue;
  followedHyperlink?: SchemeColorValue;
  /** clrScheme/`@name` (defaults to the theme name). */
  name?: string;
}

/** Font reference — a:fontRef (CT_FontReference). */
export interface FontReferenceOptions {
  /** Font collection (idx attribute: ST_FontCollectionIndex). */
  collection: "major" | "minor" | "none";
  color?: SolidFillOptions;
}

/** Shape style — a:style (CT_ShapeStyle). */
export interface DefaultShapeStyleOptions {
  lineReference: StyleMatrixReferenceOptions;
  fillReference: StyleMatrixReferenceOptions;
  effectReference: StyleMatrixReferenceOptions;
  fontReference: FontReferenceOptions;
}

/** Effect style entry — a:effectStyle (CT_EffectStyleItem). */
export interface EffectStyleOptions {
  effects?: EffectListOptions;
  scene3d?: Scene3DOptions;
  shape3d?: Shape3DOptions;
}

/** Format scheme / style matrix — a:fmtScheme (CT_StyleMatrix). */
export interface FormatSchemeOptions {
  /** Fill style list (fillStyleLst, ≥3 entries). */
  fillStyles: FillOptions[];
  /** Line style list (lnStyleLst, ≥3 entries). */
  lineStyles: OutlineOptions[];
  /** Effect style list (effectStyleLst, ≥3 entries). */
  effectStyles: EffectStyleOptions[];
  /** Background fill style list (bgFillStyleLst, ≥3 entries). */
  backgroundFillStyles: FillOptions[];
  name?: string;
}

/** Default shape/line/text definition (CT_DefaultShapeDefinition: spDef/lnDef/txDef). */
export interface DefaultShapeDefinitionOptions {
  shapeProperties?: ShapePropertiesOptions;
  bodyProperties?: BodyPropertiesOptions;
  listStyle?: TextListStyleOptions;
  shapeStyle?: DefaultShapeStyleOptions;
}

/** Object defaults — a:objectDefaults (CT_ObjectStyleDefaults). */
export interface ObjectDefaultsOptions {
  /** Shape default (spDef). */
  shapeDefault?: DefaultShapeDefinitionOptions;
  /** Line/connector default (lnDef). */
  lineDefault?: DefaultShapeDefinitionOptions;
  /** Text default (txDef). */
  textDefault?: DefaultShapeDefinitionOptions;
}

/** Theme color slot referenced by CT_ColorMapping. */
export type ColorSchemeIndex =
  | "dark1"
  | "light1"
  | "dark2"
  | "light2"
  | "accent1"
  | "accent2"
  | "accent3"
  | "accent4"
  | "accent5"
  | "accent6"
  | "hyperlink"
  | "followedHyperlink";

/** Color mapping — 12 scheme-slot remappings (CT_ColorMapping: clrMap). */
export interface ColorMappingOptions {
  background1: ColorSchemeIndex;
  text1: ColorSchemeIndex;
  background2: ColorSchemeIndex;
  text2: ColorSchemeIndex;
  accent1: ColorSchemeIndex;
  accent2: ColorSchemeIndex;
  accent3: ColorSchemeIndex;
  accent4: ColorSchemeIndex;
  accent5: ColorSchemeIndex;
  accent6: ColorSchemeIndex;
  hyperlink: ColorSchemeIndex;
  followedHyperlink: ColorSchemeIndex;
}

/** Extra color scheme (CT_ExtraColorScheme). */
export interface ExtraColorSchemeOptions {
  colorScheme: ColorSchemeOptions;
  colorMapping?: Partial<ColorMappingOptions>;
}

/** Custom color entry (CT_CustomColor) — a named color in the theme's palette. */
export interface CustomColorOptions {
  /** Palette display name (CT_CustomColor `@name`, default "") */
  name?: string;
  /** Color value (EG_ColorChoice) */
  color: SolidFillOptions;
}

/** Theme customization options (a:theme). */
export interface ThemeOptions {
  name?: string;
  colorScheme?: ColorSchemeOptions;
  fontScheme?: FontSchemeOptions;
  formatScheme?: FormatSchemeOptions;
  objectDefaults?: ObjectDefaultsOptions;
  extraColorSchemes?: ExtraColorSchemeOptions[];
  customColors?: CustomColorOptions[];
}

/** Theme override (a:themeOverride / CT_BaseStylesOverride) — a per-part subset of the theme. */
export type ThemeOverrideOptions = Pick<
  ThemeOptions,
  "colorScheme" | "fontScheme" | "formatScheme"
>;
