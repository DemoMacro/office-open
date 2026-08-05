/**
 * Paragraph options type for PPTX text paragraphs.
 *
 * @module
 */
import type { ParagraphPropertiesOptions } from "./paragraph-properties";
import type { RunOptions } from "./run";
import type { TextFieldOptions } from "./text-field";

export type ParagraphChild = RunOptions | TextFieldOptions | string;

export interface ParagraphOptions {
  /** Simple text content for the paragraph. Creates a single TextRun. */
  text?: string;
  properties?: ParagraphPropertiesOptions;
  children?: ParagraphChild[];
}
