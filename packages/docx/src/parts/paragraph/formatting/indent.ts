/**
 * Paragraph indentation module for WordprocessingML documents.
 *
 * This module provides indentation options for paragraphs including left, right,
 * hanging, and first line indentation.
 *
 * Reference: http://officeopenxml.com/WPindentation.php
 *
 * @module
 */
import type { PositiveUniversalMeasure, UniversalMeasure } from "@office-open/core";

/**
 * Properties for configuring paragraph indentation.
 *
 * Values can be specified as numbers (in twips) or as universal measures (e.g., "1in", "2.5cm").
 */
export interface IndentAttributesProperties {
  start?: number | UniversalMeasure;
  startChars?: number;
  end?: number | UniversalMeasure;
  endChars?: number;
  left?: number | UniversalMeasure;
  leftChars?: number;
  right?: number | UniversalMeasure;
  rightChars?: number;
  hanging?: number | PositiveUniversalMeasure;
  hangingChars?: number;
  firstLine?: number | PositiveUniversalMeasure;
  firstLineChars?: number;
}
