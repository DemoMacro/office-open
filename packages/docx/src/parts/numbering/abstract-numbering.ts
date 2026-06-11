/**
 * Abstract numbering definitions module for WordprocessingML documents.
 *
 * Abstract numbering definitions contain the formatting and style information
 * that can be shared across multiple numbering instances.
 *
 * Reference: http://officeopenxml.com/WPnumbering.php
 *
 * @module
 */
/**
 * Options for abstract numbering definitions.
 */
export interface AbstractNumberingOptions {
  /** Unique hex identifier for this numbering set */
  nsid?: string;
  /** Template hex identifier */
  tmpl?: string;
  /** Name of the numbering definition */
  name?: string;
  /** Paragraph style that references this numbering */
  styleLink?: string;
  /** Numbering style to inherit from */
  numStyleLink?: string;
}
