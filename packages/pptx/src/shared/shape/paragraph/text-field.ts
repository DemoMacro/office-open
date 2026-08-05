/**
 * Text field options type for PPTX text fields (a:fld).
 *
 * Mirrors CT_TextField — a dynamic field within a paragraph (date, slide number,
 * etc.), sibling to a:r runs.
 *
 * @module
 */
import type { RunPropertiesOptions } from "./run-properties";

export interface TextFieldOptions {
  /** a:fld @type — field type token (e.g. "datetimeFigureOut", "slidenum"). */
  type: string;
  /** a:fld @id — GUID identifier. */
  id?: string;
  /** a:t — display text (often a placeholder such as "‹#›" or "1/27/13"). */
  text?: string;
  /** a:rPr — run properties. */
  properties?: RunPropertiesOptions;
}
