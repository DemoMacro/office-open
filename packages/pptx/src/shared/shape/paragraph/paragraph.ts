/**
 * Paragraph options type for PPTX text paragraphs.
 *
 * @module
 */
import type { ParagraphPropertiesOptions } from "./paragraph-properties";
import type { RunOptions } from "./run";

export interface ParagraphOptions {
  /** Simple text content for the paragraph. Creates a single TextRun. */
  text?: string;
  properties?: ParagraphPropertiesOptions;
  children?: (RunOptions | string)[];
}
