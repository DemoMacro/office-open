/**
 * Paragraph border module for WordprocessingML documents.
 *
 * This module provides border options for paragraphs.
 *
 * Reference: http://officeopenxml.com/WPborders.php
 *
 * @module
 */
import type { BorderOptions } from "@shared/border";

/**
 * Options for configuring paragraph borders.
 *
 * Borders can be applied to top, bottom, left, right, and between paragraphs.
 *
 * @property top - Border for the top edge of the paragraph
 * @property bottom - Border for the bottom edge of the paragraph
 * @property left - Border for the left edge of the paragraph
 * @property right - Border for the right edge of the paragraph
 * @property between - Border between consecutive paragraphs with the same border settings
 */
export interface BordersOptions {
  /** Border for the top edge of the paragraph */
  top?: BorderOptions;
  /** Border for the bottom edge of the paragraph */
  bottom?: BorderOptions;
  /** Border for the left edge of the paragraph */
  left?: BorderOptions;
  /** Border for the right edge of the paragraph */
  right?: BorderOptions;
  /** Border between consecutive paragraphs with the same border settings */
  between?: BorderOptions;
  /** Bar border (paragraph-level bar, rendered at left of paragraph) */
  bar?: BorderOptions;
}
