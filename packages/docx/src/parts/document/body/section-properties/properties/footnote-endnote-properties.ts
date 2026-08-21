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
  /** Numbering format (ST_NumberFormat): "arabic" 1,2,3…, "lowerLetter"/"upperLetter" a,b,c…, "lowerRoman"/"upperRoman" i,ii…, "chicago", "aiueo" Japanese kana, "chineseCounting" and other script-specific counters, "bullet", "none". */
  formatType?: (typeof NumberFormat)[keyof typeof NumberFormat];
  format?: string;
  numStart?: number;
  /** Restart rule: "continuous" never, "eachSect" per section, "eachPage" per page. */
  numRestart?: (typeof NumberRestartType)[keyof typeof NumberRestartType];
}

export interface FootnotePropertiesOptions extends NumberPropertiesOptions {
  /** Where footnotes print: "pageBottom", "beneathText" under the text, "sectEnd", "docEnd". */
  pos?: (typeof FootnotePositionType)[keyof typeof FootnotePositionType];
}

export interface EndnotePropertiesOptions extends NumberPropertiesOptions {
  /** Where endnotes print: "sectEnd" end of the section, "docEnd" end of the document. */
  pos?: (typeof EndnotePositionType)[keyof typeof EndnotePositionType];
}
