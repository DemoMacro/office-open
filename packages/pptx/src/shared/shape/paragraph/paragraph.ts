/**
 * Paragraph options type for PPTX text paragraphs.
 *
 * @module
 */
import type { ParagraphPropertiesOptions } from "./paragraph-properties";
import type { RunOptions } from "./run";
import type { RunPropertiesOptions } from "./run-properties";
import type { TextFieldOptions } from "./text-field";

export type ParagraphChild = RunOptions | TextFieldOptions | string;

export interface ParagraphOptions {
  /** Simple text content for the paragraph. Creates a single TextRun. */
  text?: string;
  properties?: ParagraphPropertiesOptions;
  children?: ParagraphChild[];
  /**
   * End-paragraph run properties (a:endParaRPr). Fresh paragraphs emit a
   * default lang marker; a parsed source preserves its value; false omits the
   * element so a paragraph without endParaRPr round-trips without one.
   */
  endParagraphProperties?: RunPropertiesOptions | false;
}
