/**
 * Horizontal position module for floating drawings in WordprocessingML documents.
 *
 * This module provides horizontal positioning for floating drawing objects,
 * specifying the horizontal placement relative to a base element.
 *
 * Reference: http://officeopenxml.com/drwPicFloating-position.php
 *
 * @module
 */
import { convertToEmu } from "@office-open/core";
import { element } from "@office-open/xml";

import { HorizontalPositionAlign, HorizontalPositionRelativeFrom } from "./floating-position";
import type { HorizontalPositionOptions } from "./floating-position";

/**
 * Creates a horizontal position element for floating drawings.
 *
 * The positionH element specifies the horizontal positioning of a floating
 * object relative to a base element (page, margin, column, etc.).
 *
 * Reference: https://www.datypic.com/sc/ooxml/e-wp_positionH-1.html
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_PosH">
 *   <xsd:choice>
 *     <xsd:element name="align" type="ST_AlignH"/>
 *     <xsd:element name="posOffset" type="ST_PositionOffset"/>
 *   </xsd:choice>
 *   <xsd:attribute name="relativeFrom" type="ST_RelFromH" use="required"/>
 * </xsd:complexType>
 * ```
 *
 * @param options - Horizontal position configuration
 * @returns The positionH XML string
 *
 * @example
 * ```typescript
 * // Align to the left of the page
 * createHorizontalPosition({
 *   relative: HorizontalPositionRelativeFrom.PAGE,
 *   align: HorizontalPositionAlign.LEFT,
 * });
 *
 * // Offset from the margin
 * createHorizontalPosition({
 *   relative: HorizontalPositionRelativeFrom.MARGIN,
 *   offset: 914400, // 1 inch in EMUs
 * });
 * ```
 */
export const createHorizontalPosition = ({
  relative,
  align,
  offset,
  percentOffset,
}: HorizontalPositionOptions): string => {
  const relativeFrom = relative ?? HorizontalPositionRelativeFrom.PAGE;
  if (percentOffset !== undefined) {
    const position = `<wp:positionH relativeFrom="${relativeFrom}"><wp14:pctPosHOffset>${Math.round(percentOffset * 1000)}</wp14:pctPosHOffset></wp:positionH>`;
    if (offset === undefined) return position;
    return (
      '<mc:AlternateContent><mc:Choice Requires="wp14">' +
      position +
      `</mc:Choice><mc:Fallback><wp:positionH relativeFrom="${relativeFrom}"><wp:posOffset>${convertToEmu(offset)}</wp:posOffset></wp:positionH></mc:Fallback></mc:AlternateContent>`
    );
  }

  const child = align
    ? `<wp:align>${align}</wp:align>`
    : offset !== undefined
      ? `<wp:posOffset>${convertToEmu(offset)}</wp:posOffset>`
      : `<wp:align>${HorizontalPositionAlign.LEFT}</wp:align>`;

  return element("wp:positionH", { relativeFrom }, [child]);
};
