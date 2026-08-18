/**
 * Text list styles — re-exported from core DrawingML.
 *
 * The bare CT_TextListStyle descriptor (a:lstStyle levels), the master
 * txStyles groups, and the MS Office default live in core. This module
 * re-exports them for internal callers (parse.ts, parts/index.ts,
 * slide-master.ts).
 *
 * @module
 */

export {
  textListStyleDesc,
  textStylesDesc,
  stringifyTextStyles,
  parseTextStyles,
  DEFAULT_TEXT_STYLES,
} from "@office-open/core/drawing";

export type { TextListStyleOptions, TextStylesOptions } from "@office-open/core/drawing";
