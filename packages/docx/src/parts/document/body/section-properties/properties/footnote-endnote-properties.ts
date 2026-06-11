/**
 * Footnote and endnote properties module for WordprocessingML section properties.
 *
 * Specifies footnote/endnote placement and numbering format within a section.
 *
 * Reference: ISO/IEC 29500-4, CT_FtnProps / CT_EdnProps
 *
 * @module
 */
import type { NumberFormat } from "@shared/constants";

/**
 * Footnote position types.
 *
 * @publicApi
 */
export const FootnotePositionType = {
  PAGE_BOTTOM: "pageBottom",
  BENEATH_TEXT: "beneathText",
  SECT_END: "sectEnd",
  DOC_END: "docEnd",
} as const;

/**
 * Endnote position types.
 *
 * @publicApi
 */
export const EndnotePositionType = {
  SECT_END: "sectEnd",
  DOC_END: "docEnd",
} as const;

/**
 * Number restart types for footnotes/endnotes.
 *
 * @publicApi
 */
export const NumberRestartType = {
  CONTINUOUS: "continuous",
  EACH_SECT: "eachSect",
  EACH_PAGE: "eachPage",
} as const;

interface NumberPropertiesOptions {
  formatType?: (typeof NumberFormat)[keyof typeof NumberFormat];
  format?: string;
  numStart?: number;
  numRestart?: (typeof NumberRestartType)[keyof typeof NumberRestartType];
}

export interface FootnotePropertiesOptions extends NumberPropertiesOptions {
  pos?: (typeof FootnotePositionType)[keyof typeof FootnotePositionType];
}

export interface EndnotePropertiesOptions extends NumberPropertiesOptions {
  pos?: (typeof EndnotePositionType)[keyof typeof EndnotePositionType];
}
