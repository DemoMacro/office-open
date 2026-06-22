/**
 * Paragraph spacing module for WordprocessingML documents.
 *
 * This module provides spacing options for paragraphs including space before,
 * space after, and line spacing.
 *
 * Reference: http://officeopenxml.com/WPspacing.php
 *
 * @module
 */
import type { PositiveUniversalMeasure } from "@office-open/core";

/**
 * Line spacing rule types.
 *
 * Specifies how the line height is calculated.
 *
 * @publicApi
 */
export const LineRuleType = {
  /** Line spacing is at least the specified value */
  AT_LEAST: "atLeast",
  /** Line spacing is exactly the specified value */
  EXACTLY: "exactly",
  /** Line spacing is exactly the specified value (alias for EXACTLY) */
  EXACT: "exact",
  /** Line spacing is automatically determined based on content */
  AUTO: "auto",
} as const;

/**
 * Properties for configuring paragraph spacing.
 *
 * All values are in twips (twentieths of a point) unless otherwise specified.
 */
export interface SpacingProperties {
  /** Spacing after the paragraph in twips or as a PositiveUniversalMeasure */
  after?: number | PositiveUniversalMeasure;
  /** Spacing before the paragraph in twips or as a PositiveUniversalMeasure */
  before?: number | PositiveUniversalMeasure;
  /** Line spacing value in twips or as a PositiveUniversalMeasure (depends on lineRule) */
  line?: number | PositiveUniversalMeasure;
  /** How to interpret the line spacing value */
  lineRule?: (typeof LineRuleType)[keyof typeof LineRuleType];
  /** Use automatic spacing before the paragraph */
  beforeAutoSpacing?: boolean;
  /** Use automatic spacing after the paragraph */
  afterAutoSpacing?: boolean;
  /** Spacing before the paragraph in line units */
  beforeLines?: number;
  /** Spacing after the paragraph in line units */
  afterLines?: number;
}
