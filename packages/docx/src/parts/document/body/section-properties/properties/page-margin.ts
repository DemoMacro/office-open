/**
 * Page margin module for WordprocessingML section properties.
 *
 * Defines page margins for document sections.
 *
 * Reference: http://officeopenxml.com/WPsectionPr.php
 *
 * @module
 */
import { signedTwipsMeasureValue, twipsMeasureValue } from "@office-open/core";
import type { PositiveUniversalMeasure, UniversalMeasure } from "@office-open/core";

/**
 * Options for configuring page margins.
 *
 * All measurements are in twips (1/20th of a point) or universal measure.
 *
 * @property top - Top margin
 * @property right - Right margin
 * @property bottom - Bottom margin
 * @property left - Left margin
 * @property header - Header margin (distance from top of page to header)
 * @property footer - Footer margin (distance from bottom of page to footer)
 * @property gutter - Gutter margin for binding
 */
export interface PageMarginAttributes {
  /** Top margin in twips or universal measure */
  top?: number | UniversalMeasure;
  /** Right margin in twips or universal measure */
  right?: number | PositiveUniversalMeasure;
  /** Bottom margin in twips or universal measure */
  bottom?: number | UniversalMeasure;
  /** Left margin in twips or universal measure */
  left?: number | PositiveUniversalMeasure;
  /** Header margin (distance from top of page to header) in twips or universal measure */
  header?: number | PositiveUniversalMeasure;
  /** Footer margin (distance from bottom of page to footer) in twips or universal measure */
  footer?: number | PositiveUniversalMeasure;
  /** Gutter margin for binding in twips or universal measure */
  gutter?: number | PositiveUniversalMeasure;
}

/**
 * Creates page margins (pgMar) for a document section.
 *
 * This element specifies the page margins for all pages in a section,
 * including top, bottom, left, right, header, footer, and gutter margins.
 *
 * Reference: http://officeopenxml.com/WPsectionPr.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_PageMar">
 *   <xsd:attribute name="top" type="ST_SignedTwipsMeasure" use="required"/>
 *   <xsd:attribute name="right" type="s:ST_TwipsMeasure" use="required"/>
 *   <xsd:attribute name="bottom" type="ST_SignedTwipsMeasure" use="required"/>
 *   <xsd:attribute name="left" type="s:ST_TwipsMeasure" use="required"/>
 *   <xsd:attribute name="header" type="s:ST_TwipsMeasure" use="required"/>
 *   <xsd:attribute name="footer" type="s:ST_TwipsMeasure" use="required"/>
 *   <xsd:attribute name="gutter" type="s:ST_TwipsMeasure" use="required"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Create page margins with 1 inch margins (1440 twips = 1 inch)
 * createPageMargin(1440, 1440, 1440, 1440, 720, 720, 0);
 * ```
 */
export const createPageMargin = (
  top: number | UniversalMeasure,
  right: number | PositiveUniversalMeasure,
  bottom: number | UniversalMeasure,
  left: number | PositiveUniversalMeasure,
  header: number | PositiveUniversalMeasure,
  footer: number | PositiveUniversalMeasure,
  gutter: number | PositiveUniversalMeasure,
): string =>
  `<w:pgMar w:bottom="${signedTwipsMeasureValue(bottom)}" w:footer="${twipsMeasureValue(footer)}" w:gutter="${twipsMeasureValue(gutter)}" w:header="${twipsMeasureValue(header)}" w:left="${twipsMeasureValue(left)}" w:right="${twipsMeasureValue(right)}" w:top="${signedTwipsMeasureValue(top)}"/>`;
