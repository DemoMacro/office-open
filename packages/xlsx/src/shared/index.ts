/**
 * Shared module — cross-part types for XLSX generation.
 *
 * @module
 */
export { Media, type MediaData } from "./media";
export { createThemeXml } from "./theme";
export { lintWorkbookFormulas, type FormulaIssue } from "./formula-lint";
export { colorAttributes, parseColorOptions, type ColorOptions } from "./color";
