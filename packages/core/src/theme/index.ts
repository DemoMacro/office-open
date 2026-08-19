/**
 * Theme module — unified theme generation and round-trip for OOXML documents.
 *
 * @module
 */
export { createThemeXml } from "./default-theme";
export { buildThemeXml } from "./build-theme-xml";
export { DEFAULT_COLORS } from "./default-colors";
export { themeDesc } from "./theme-descriptors";
export { themeManagerDesc, themeOverrideDesc } from "./theme-override";
export type { ThemeManagerOptions } from "./theme-override";
export { DEFAULT_COLOR_MAPPING, parseColorMapping, stringifyColorMapping } from "./color-mapping";
export { parseShapeStyle, stringifyShapeStyle } from "./style-matrix";
export type {
  ColorMappingOptions,
  ColorSchemeIndex,
  ColorSchemeOptions,
  CustomColorOptions,
  DefaultShapeDefinitionOptions,
  EffectStyleOptions,
  ExtraColorSchemeOptions,
  FontCollectionOptions,
  FontReferenceOptions,
  FontSchemeOptions,
  FormatSchemeOptions,
  ObjectDefaultsOptions,
  DefaultShapeStyleOptions,
  StyleMatrixReferenceOptions,
  SupplementalFontOptions,
  TextFontOptions,
  ThemeOptions,
  ThemeOverrideOptions,
} from "./theme-options";
