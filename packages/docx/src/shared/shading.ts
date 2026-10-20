/**
 * Shading module for WordprocessingML documents.
 *
 * Shading is used to apply background colors and patterns to paragraphs,
 * table cells, and text runs. The shading type is identical in all places.
 *
 * Reference: http://officeopenxml.com/WPshading.php
 *
 * @see http://officeopenxml.com/WPtableShading.php
 * @see http://officeopenxml.com/WPtableCellProperties-Shading.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Shd">
 *   <xsd:attribute name="val" type="ST_Shd" use="required"/>
 *   <xsd:attribute name="color" type="ST_HexColor" use="optional"/>
 *   <xsd:attribute name="themeColor" type="ST_ThemeColor" use="optional"/>
 *   <xsd:attribute name="themeTint" type="ST_UcharHexNumber" use="optional"/>
 *   <xsd:attribute name="themeShade" type="ST_UcharHexNumber" use="optional"/>
 *   <xsd:attribute name="fill" type="ST_HexColor" use="optional"/>
 *   <xsd:attribute name="themeFill" type="ST_ThemeColor" use="optional"/>
 *   <xsd:attribute name="themeFillTint" type="ST_UcharHexNumber" use="optional"/>
 *   <xsd:attribute name="themeFillShade" type="ST_UcharHexNumber" use="optional"/>
 * </xsd:complexType>
 * ```
 *
 * @module
 */
import { xsdShadingPattern } from "@office-open/core";
import type { ThemeColor } from "@office-open/core";
import type { HexColorOrAuto, UcharHexNumber } from "@office-open/core";
import { attr } from "@office-open/xml";
import type { Element } from "@office-open/xml";

/**
 * Properties for configuring shading.
 *
 * @property fill - Background fill color in hex format (e.g., "FF0000" for red)
 * @property color - Pattern color in hex format
 * @property type - Shading pattern type
 */
export interface ShadingProperties {
  fill?: HexColorOrAuto;
  color?: HexColorOrAuto;
  /** Source `@w:fill` when it is outside the canonical hex/auto surface; round-trip only. */
  fillRaw?: string;
  /** Source `@w:color` when it is outside the canonical hex/auto surface; round-trip only. */
  colorRaw?: string;
  /** Fill pattern (ST_Shd): "clear" solid fill, "nil" no fill, "percentN" dot-density percent, "horizontalStripe"/"diagonalStripe"/"horizontalCross"/"diagonalCross" hatch patterns. */
  type?: (typeof ShadingType)[keyof typeof ShadingType] | null;
  /** Theme color reference */
  themeColor?: ThemeColor;
  /** Theme color tint (2-char hex) */
  themeTint?: UcharHexNumber;
  /** Source `@w:themeTint` when it is not 2-char hex; round-trip only. */
  themeTintRaw?: string;
  /** Theme color shade (2-char hex) */
  themeShade?: UcharHexNumber;
  /** Source `@w:themeShade` when it is not 2-char hex; round-trip only. */
  themeShadeRaw?: string;
  /** Theme fill color reference */
  themeFill?: ThemeColor;
  /** Theme fill tint (2-char hex) */
  themeFillTint?: UcharHexNumber;
  /** Source `@w:themeFillTint` when it is not 2-char hex; round-trip only. */
  themeFillTintRaw?: string;
  /** Theme fill shade (2-char hex) */
  themeFillShade?: UcharHexNumber;
  /** Source `@w:themeFillShade` when it is not 2-char hex; round-trip only. */
  themeFillShadeRaw?: string;
}

/**
 * Shading pattern types.
 *
 * Specifies the pattern used for shading. The pattern combines the fill
 * color and the pattern color.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_Shd">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="nil"/>
 *     <xsd:enumeration value="clear"/>
 *     <xsd:enumeration value="solid"/>
 *     <xsd:enumeration value="horzStripe"/>
 *     <xsd:enumeration value="vertStripe"/>
 *     <xsd:enumeration value="reverseDiagStripe"/>
 *     <xsd:enumeration value="diagStripe"/>
 *     <xsd:enumeration value="horzCross"/>
 *     <xsd:enumeration value="diagCross"/>
 *     <!-- ... percent values ... -->
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @publicApi
 */
export const ShadingType = {
  /** Clear shading - no pattern, fill color only */
  CLEAR: "clear",
  DIAGONAL_CROSS: "diagonalCross",
  DIAGONAL_STRIPE: "diagonalStripe",
  HORIZONTAL_CROSS: "horizontalCross",
  HORIZONTAL_STRIPE: "horizontalStripe",
  NIL: "nil",
  PERCENT_10: "percent10",
  PERCENT_12: "percent12",
  PERCENT_15: "percent15",
  PERCENT_20: "percent20",
  PERCENT_25: "percent25",
  PERCENT_30: "percent30",
  PERCENT_35: "percent35",
  PERCENT_37: "percent37",
  PERCENT_40: "percent40",
  PERCENT_45: "percent45",
  PERCENT_5: "percent5",
  PERCENT_50: "percent50",
  PERCENT_55: "percent55",
  PERCENT_60: "percent60",
  PERCENT_62: "percent62",
  PERCENT_65: "percent65",
  PERCENT_70: "percent70",
  PERCENT_75: "percent75",
  PERCENT_80: "percent80",
  PERCENT_85: "percent85",
  PERCENT_87: "percent87",
  PERCENT_90: "percent90",
  PERCENT_95: "percent95",
  REVERSE_DIAGONAL_STRIPE: "reverseDiagonalStripe",
  SOLID: "solid",
  THIN_DIAGONAL_CROSS: "thinDiagonalCross",
  THIN_DIAGONAL_STRIPE: "thinDiagonalStripe",
  THIN_HORIZONTAL_CROSS: "thinHorizontalCross",
  THIN_HORIZONTAL_STRIPE: "thinHorizontalStripe",
  THIN_REVERSE_DIAGONAL_STRIPE: "thinReverseDiagonalStripe",
  THIN_VERTICAL_STRIPE: "thinVerticalStripe",
  VERTICAL_STRIPE: "verticalStripe",
} as const;

/**
 * Parse a w:shd (CT_Shd) element into ShadingProperties.
 *
 * Reads every CT_Shd attribute (fill/color/val plus the theme* family), so the
 * result round-trips losslessly — paragraph, table-cell, and run shading all
 * share this single reader. Returns undefined when the element carries no data.
 */
export function parseShading(shd: Element): ShadingProperties | undefined {
  const shading: ShadingProperties = {};
  const fill = attr(shd, "w:fill");
  if (fill) {
    shading.fill = fill;
    if (!/^(?:auto|[0-9a-fA-F]{6})$/.test(fill)) shading.fillRaw = fill;
  }
  const color = attr(shd, "w:color");
  if (color) {
    shading.color = color;
    if (!/^(?:auto|[0-9a-fA-F]{6})$/.test(color)) shading.colorRaw = color;
  }
  const val = attr(shd, "w:val");
  shading.type =
    val === undefined ? null : (xsdShadingPattern.from(val) as ShadingProperties["type"]);
  const themeColor = attr(shd, "w:themeColor");
  if (themeColor) {
    shading.themeColor = themeColor as ThemeColor;
  }
  const themeTint = attr(shd, "w:themeTint");
  if (themeTint) {
    if (/^[0-9a-fA-F]{2}$/.test(themeTint)) shading.themeTint = themeTint;
    else shading.themeTintRaw = themeTint;
  }
  const themeShade = attr(shd, "w:themeShade");
  if (themeShade) {
    if (/^[0-9a-fA-F]{2}$/.test(themeShade)) shading.themeShade = themeShade;
    else shading.themeShadeRaw = themeShade;
  }
  const themeFill = attr(shd, "w:themeFill");
  if (themeFill) {
    shading.themeFill = themeFill as ThemeColor;
  }
  const themeFillTint = attr(shd, "w:themeFillTint");
  if (themeFillTint) {
    if (/^[0-9a-fA-F]{2}$/.test(themeFillTint)) shading.themeFillTint = themeFillTint;
    else shading.themeFillTintRaw = themeFillTint;
  }
  const themeFillShade = attr(shd, "w:themeFillShade");
  if (themeFillShade) {
    if (/^[0-9a-fA-F]{2}$/.test(themeFillShade)) shading.themeFillShade = themeFillShade;
    else shading.themeFillShadeRaw = themeFillShade;
  }
  return shading;
}
