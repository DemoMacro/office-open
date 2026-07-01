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
import { ThemeColor } from "@office-open/core";
import { attr } from "@office-open/xml";
import type { Element } from "@office-open/xml";

/**
 * Properties for configuring shading.
 *
 * @property fill - Background fill color in hex format (e.g., "FF0000" for red)
 * @property color - Pattern color in hex format
 * @property type - Shading pattern type
 */
export interface ShadingAttributesProperties {
  fill?: string;
  color?: string;
  type?: (typeof ShadingType)[keyof typeof ShadingType];
  /** Theme color reference */
  themeColor?: (typeof ThemeColor)[keyof typeof ThemeColor];
  /** Theme color tint (2-char hex) */
  themeTint?: string;
  /** Theme color shade (2-char hex) */
  themeShade?: string;
  /** Theme fill color reference */
  themeFill?: (typeof ThemeColor)[keyof typeof ThemeColor];
  /** Theme fill tint (2-char hex) */
  themeFillTint?: string;
  /** Theme fill shade (2-char hex) */
  themeFillShade?: string;
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
  DIAGONAL_CROSS: "diagCross",
  DIAGONAL_STRIPE: "diagStripe",
  HORIZONTAL_CROSS: "horzCross",
  HORIZONTAL_STRIPE: "horzStripe",
  NIL: "nil",
  PERCENT_10: "pct10",
  PERCENT_12: "pct12",
  PERCENT_15: "pct15",
  PERCENT_20: "pct20",
  PERCENT_25: "pct25",
  PERCENT_30: "pct30",
  PERCENT_35: "pct35",
  PERCENT_37: "pct37",
  PERCENT_40: "pct40",
  PERCENT_45: "pct45",
  PERCENT_5: "pct5",
  PERCENT_50: "pct50",
  PERCENT_55: "pct55",
  PERCENT_60: "pct60",
  PERCENT_62: "pct62",
  PERCENT_65: "pct65",
  PERCENT_70: "pct70",
  PERCENT_75: "pct75",
  PERCENT_80: "pct80",
  PERCENT_85: "pct85",
  PERCENT_87: "pct87",
  PERCENT_90: "pct90",
  PERCENT_95: "pct95",
  REVERSE_DIAGONAL_STRIPE: "reverseDiagStripe",
  SOLID: "solid",
  THIN_DIAGONAL_CROSS: "thinDiagCross",
  THIN_DIAGONAL_STRIPE: "thinDiagStripe",
  THIN_HORIZONTAL_CROSS: "thinHorzCross",
  THIN_REVERSE_DIAGONAL_STRIPE: "thinReverseDiagStripe",
  THIN_VERTICAL_STRIPE: "thinVertStripe",
  VERTICAL_STRIPE: "vertStripe",
} as const;

const THEME_COLORS = Object.values(ThemeColor) as readonly string[];

/**
 * Parse a w:shd (CT_Shd) element into ShadingAttributesProperties.
 *
 * Reads every CT_Shd attribute (fill/color/val plus the theme* family), so the
 * result round-trips losslessly — paragraph, table-cell, and run shading all
 * share this single reader. Returns undefined when the element carries no data.
 */
export function parseShading(shd: Element): ShadingAttributesProperties | undefined {
  const shading: ShadingAttributesProperties = {};
  const fill = attr(shd, "w:fill");
  if (fill) shading.fill = fill;
  const color = attr(shd, "w:color");
  if (color) shading.color = color;
  const val = attr(shd, "w:val");
  if (val) shading.type = val as ShadingAttributesProperties["type"];
  const themeColor = attr(shd, "w:themeColor");
  if (themeColor && THEME_COLORS.includes(themeColor)) {
    shading.themeColor = themeColor as ShadingAttributesProperties["themeColor"];
  }
  const themeTint = attr(shd, "w:themeTint");
  if (themeTint) shading.themeTint = themeTint;
  const themeShade = attr(shd, "w:themeShade");
  if (themeShade) shading.themeShade = themeShade;
  const themeFill = attr(shd, "w:themeFill");
  if (themeFill && THEME_COLORS.includes(themeFill)) {
    shading.themeFill = themeFill as ShadingAttributesProperties["themeFill"];
  }
  const themeFillTint = attr(shd, "w:themeFillTint");
  if (themeFillTint) shading.themeFillTint = themeFillTint;
  const themeFillShade = attr(shd, "w:themeFillShade");
  if (themeFillShade) shading.themeFillShade = themeFillShade;
  if (Object.keys(shading).length === 0) return undefined;
  return shading;
}
