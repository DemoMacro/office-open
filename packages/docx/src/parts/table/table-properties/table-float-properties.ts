/**
 * Table float properties module for WordprocessingML documents.
 *
 * This module provides floating table positioning options, allowing tables
 * to float with text wrapping around them.
 *
 * Reference: http://officeopenxml.com/WPtableFloating.php
 *
 * @module
 */

/**
 * Anchor types for floating table positioning.
 *
 * Specifies the base object from which positioning is determined.
 */
export const TableAnchorType = {
  MARGIN: "margin",
  PAGE: "page",
  TEXT: "text",
} as const;

/**
 * Relative horizontal position values for floating tables.
 */
export const RelativeHorizontalPosition = {
  CENTER: "center",
  INSIDE: "inside",
  LEFT: "left",
  OUTSIDE: "outside",
  RIGHT: "right",
} as const;

/**
 * Relative vertical position values for floating tables.
 */
export const RelativeVerticalPosition = {
  BOTTOM: "bottom",
  CENTER: "center",
  INLINE: "inline",
  INSIDE: "inside",
  OUTSIDE: "outside",
  TOP: "top",
} as const;

/**
 * Table overlap behavior types.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_TblOverlap">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="never"/>
 *     <xsd:enumeration value="overlap"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 */
export const OverlapType = {
  NEVER: "never",
  OVERLAP: "overlap",
} as const;

export interface TableFloatOptions {
  /* CSpell:disable */
  /** Reference edge for the horizontal position: margin, page, or text column. */
  horizontalAnchor?: (typeof TableAnchorType)[keyof typeof TableAnchorType];
  absoluteHorizontalPosition?: number;
  /** Horizontal alignment within the anchor: "inside"/"outside" the binding (odd/even page side), or an edge. */
  relativeHorizontalPosition?: (typeof RelativeHorizontalPosition)[keyof typeof RelativeHorizontalPosition];
  /** Reference edge for the vertical position: margin, page, or text paragraph. */
  verticalAnchor?: (typeof TableAnchorType)[keyof typeof TableAnchorType];
  absoluteVerticalPosition?: number;
  /** Vertical alignment within the anchor; "inline" flows in the text line. */
  relativeVerticalPosition?: (typeof RelativeVerticalPosition)[keyof typeof RelativeVerticalPosition];
  bottomFromText?: number;
  topFromText?: number;
  leftFromText?: number;
  rightFromText?: number;
  /** Whether two floating tables may overlap. */
  overlap?: (typeof OverlapType)[keyof typeof OverlapType];
  /* CSpell:enable */
}
