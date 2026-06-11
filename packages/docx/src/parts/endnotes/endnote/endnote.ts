/**
 * Endnote module for WordprocessingML documents.
 *
 * @module
 */
import type { ParagraphOptions } from "@parts/paragraph/paragraph";

export const EndnoteType = {
  CONTINUATION_SEPARATOR: "continuationSeparator",

  SEPARATOR: "separator",
} as const;

export interface EndnoteOptions {
  id: number;
  type?: (typeof EndnoteType)[keyof typeof EndnoteType];
  children: (ParagraphOptions | string)[];
}
