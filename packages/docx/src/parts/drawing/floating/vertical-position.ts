/**
 * Vertical position module for floating drawings in WordprocessingML documents.
 *
 * This module provides vertical positioning for floating drawing objects,
 * specifying the vertical placement relative to a base element.
 *
 * Reference: http://officeopenxml.com/drwPicFloating-position.php
 *
 * @module
 */
import { element } from "@office-open/xml";

import { VerticalPositionAlign, VerticalPositionRelativeFrom } from "./floating-position";
import type { VerticalPositionOptions } from "./floating-position";

/**
 * Creates a vertical position element for floating drawings.
 *
 * The positionV element specifies the vertical positioning of a floating
 * object relative to a base element (page, margin, paragraph, line, etc.).
 *
 * Reference: https://www.datypic.com/sc/ooxml/e-wp_positionV-1.html
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_PosV">
 *   <xsd:choice>
 *     <xsd:element name="align" type="ST_AlignV"/>
 *     <xsd:element name="posOffset" type="ST_PositionOffset"/>
 *   </xsd:choice>
 *   <xsd:attribute name="relativeFrom" type="ST_RelFromV" use="required"/>
 * </xsd:complexType>
 * ```
 *
 * @param options - Vertical position configuration
 * @returns The positionV XML string
 *
 * @example
 * ```typescript
 * // Align to the top of the page
 * createVerticalPosition({
 *   relative: VerticalPositionRelativeFrom.PAGE,
 *   align: VerticalPositionAlign.TOP,
 * });
 *
 * // Offset from the paragraph
 * createVerticalPosition({
 *   relative: VerticalPositionRelativeFrom.PARAGRAPH,
 *   offset: 457200, // 0.5 inch in EMUs
 * });
 * ```
 */
export const createVerticalPosition = ({
  relative,
  align,
  offset,
}: VerticalPositionOptions): string => {
  const child = align
    ? `<wp:align>${align}</wp:align>`
    : offset !== undefined
      ? `<wp:posOffset>${offset}</wp:posOffset>`
      : `<wp:align>${VerticalPositionAlign.TOP}</wp:align>`;

  return element(
    "wp:positionV",
    {
      relativeFrom: relative ?? VerticalPositionRelativeFrom.PAGE,
    },
    [child],
  );
};
