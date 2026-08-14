/**
 * Theme module — unified theme generation and round-trip for OOXML documents.
 *
 * @module
 */
export { createThemeXml } from "./default-theme";
export { buildThemeXml } from "./build-theme-xml";
export { DEFAULT_COLORS } from "./default-colors";
export { themeDesc } from "./theme-descriptors";
export { DEFAULT_COLOR_MAPPING, parseColorMapping, stringifyColorMapping } from "./color-mapping";
export type {
  ColorMappingOptions,
  ColorSchemeIndex,
  ColorSchemeOptions,
  DefaultShapeDefinitionOptions,
  EffectStyleOptions,
  ExtraColorSchemeOptions,
  FontCollectionOptions,
  FontReferenceOptions,
  FontSchemeOptions,
  FormatSchemeOptions,
  ObjectDefaultsOptions,
  ShapeStyleOptions,
  StyleMatrixReferenceOptions,
  SupplementalFontOptions,
  TextFontOptions,
  ThemeOptions,
} from "./theme-options";
