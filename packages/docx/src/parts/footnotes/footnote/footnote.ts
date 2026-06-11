/**
 * Footnote module for WordprocessingML documents.
 *
 * This module provides support for footnotes that appear at the bottom
 * of the page referenced from the main document text.
 *
 * Reference: http://officeopenxml.com/WPfootnotes.php
 *
 * @module
 */
import type { ParagraphOptions } from "@parts/paragraph/paragraph";

/**
 * Enumeration of footnote types.
 *
 * Reference: http://officeopenxml.com/WPfootnotes.php
 *
 * @publicApi
 */
export const FootnoteType = {
  /** Separator line between body text and footnotes */
  SEPERATOR: "separator",
  /** Continuation separator for footnotes spanning pages */
  CONTINUATION_SEPERATOR: "continuationSeparator",
} as const;

/**
 * Options for creating a Footnote.
 *
 * @property id - Unique numeric identifier for this footnote
 * @property type - Type of footnote (separator, continuationSeparator, or normal)
 * @property children - Array of paragraphs that make up the footnote content
 *
 * @see {@link Footnote}
 */
export interface FootnoteOptions {
  /** Unique numeric identifier for this footnote */
  id: number;
  /** Type of footnote (separator or continuation separator) */
  type?: (typeof FootnoteType)[keyof typeof FootnoteType];
  /** Content of the footnote (paragraphs, tables, etc.) */
  children: (ParagraphOptions | string)[];
}
