/**
 * Numbered item reference module for WordprocessingML documents.
 *
 * This module provides cross-references to numbered items (such as
 * numbered paragraphs, list items, or headings) within a document.
 *
 * Reference: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/7088a8ce-e784-49d4-94b8-cba6ef8fce78
 *
 * @module
 */

/**
 * Format options for numbered item references.
 *
 * Specifies how the paragraph number should be displayed when referenced.
 */
export enum NumberedItemReferenceFormat {
  NONE = "none",
  /**
   * \r option - inserts the paragraph number of the bookmarked paragraph in relative context, or relative to its position in the numbering scheme
   */
  RELATIVE = "relative",
  /**
   * \n option - causes the field result to be the paragraph number without trailing periods. No information about prior numbered levels is displayed unless it is included as part of the current level.
   */
  NO_CONTEXT = "no_context",
  /**
   * \w option - causes the field result to be the entire paragraph number without trailing periods, regardless of the location of the REF field.
   */
  FULL_CONTEXT = "full_context",
}

export interface NumberedItemReferenceOptions {
  /**
   * \h option - Creates a hyperlink to the bookmarked paragraph.
   * @default true
   */
  hyperlink?: boolean;
  /**
   * Which switch to use for the reference format
   * @default NumberedItemReferenceFormat.FULL_CONTEXT
   */
  referenceFormat?: NumberedItemReferenceFormat;
}
