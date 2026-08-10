/**
 * Text list style (CT_TextListStyle) — re-exported from core DrawingML.
 *
 * The list-style descriptor, MS Office default, and stringify/parse helpers
 * live in core. This module re-exports them for backward compatibility with
 * internal callers (parse.ts, parts/index.ts, slide-master.ts).
 *
 * @module
 */

export {
  textListStyleDesc,
  stringifyTextListStyle,
  parseTextListStyle,
  DEFAULT_TEXT_LIST_STYLE,
} from "@office-open/core/drawingml";

export type {
  TextListStyleOptions,
  TextListStyleGroupOptions,
  TextListStyleLevelOptions,
  TextListStyleBulletOptions,
  TextListStyleRunOptions,
} from "@office-open/core/drawingml";
