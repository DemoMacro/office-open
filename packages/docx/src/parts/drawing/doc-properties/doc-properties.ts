/**
 * Document Properties module for DrawingML elements.
 *
 * This module provides non-visual properties for drawing elements,
 * including name, description, accessibility information, and hyperlinks.
 *
 * Reference: https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_docPr_topic_ID0ES32OB.html
 *
 * @module
 */

/**
 * Options for hyperlinks on a drawing element.
 */
export interface HyperlinkOptions {
  /** URL for click hyperlink */
  click?: string;
  /** URL for hover hyperlink */
  hover?: string;
}

/**
 * Options for configuring document properties of a drawing.
 *
 * @see {@link DocProperties}
 */
export interface DocPropertiesOptions {
  /** Name of the drawing element (used for identification) */
  name: string;
  /** Description/alt text for accessibility */
  description?: string;
  /** Title of the drawing element */
  title?: string;
  id?: string;
  /** Hyperlink options for click and hover actions */
  hyperlink?: HyperlinkOptions;
}
