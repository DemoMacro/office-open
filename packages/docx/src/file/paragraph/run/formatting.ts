/**
 * Run formatting module for WordprocessingML documents.
 *
 * This module provides character-level formatting elements including
 * spacing, color, and highlighting.
 *
 * Reference: http://officeopenxml.com/WPtextFormatting.php
 *
 * @module
 */
import { XmlComponent, attrObj } from "@file/xml-components";
import type { IXmlableObject } from "@file/xml-components";
import { hexColorValue, signedTwipsMeasureValue, uCharHexNumber } from "@util/values";
import type { ThemeColor, UniversalMeasure } from "@util/values";

/**
 * Options for theme color configuration.
 *
 * @property val - Explicit hex color value (e.g., "FF0000" or "auto")
 * @property themeColor - Theme color reference
 * @property themeTint - Theme color tint (2-char hex, e.g., "99")
 * @property themeShade - Theme color shade (2-char hex, e.g., "BF")
 */
export interface ColorOptions {
  val?: string;
  themeColor?: (typeof ThemeColor)[keyof typeof ThemeColor];
  themeTint?: string;
  themeShade?: string;
}

/**
 * Represents character spacing (tracking) in a run.
 *
 * Adjusts the space between characters in the text.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_SignedTwipsMeasure">
 *   <xsd:attribute name="val" type="ST_SignedTwipsMeasure" use="required"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * new CharacterSpacing(20); // 20 twips (1 point)
 * new CharacterSpacing("0.5pt"); // Half point
 * ```
 *
 * @internal
 */
export class CharacterSpacing extends XmlComponent {
  public constructor(value: number | UniversalMeasure) {
    super("w:spacing");
    this.root.push({ _attr: { "w:val": signedTwipsMeasureValue(value) } });
  }
}

/**
 * Represents text color in a run.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Color">
 *   <xsd:attribute name="val" type="ST_HexColor" use="required"/>
 *   <xsd:attribute name="themeColor" type="ST_ThemeColor" use="optional"/>
 *   <xsd:attribute name="themeTint" type="ST_UcharHexNumber" use="optional"/>
 *   <xsd:attribute name="themeShade" type="ST_UcharHexNumber" use="optional"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * new Color("FF0000"); // Red text (backward compatible)
 * new Color("auto"); // Automatic color
 * new Color({ themeColor: ThemeColor.ACCENT1, themeTint: "99" }); // Theme color with tint
 * ```
 *
 * @internal
 */
export class Color extends XmlComponent {
  public constructor(colorOrOptions: string | ColorOptions) {
    super("w:color");

    if (typeof colorOrOptions === "string") {
      this.root.push({ _attr: { "w:val": hexColorValue(colorOrOptions) } });
      return;
    }

    const opts = colorOrOptions;
    const attrs: Record<string, string> = {};
    if (opts.val !== undefined) {
      attrs["w:val"] = hexColorValue(opts.val);
    }
    if (opts.themeColor !== undefined) {
      attrs["w:themeColor"] = opts.themeColor;
    }
    if (opts.themeTint !== undefined) {
      attrs["w:themeTint"] = uCharHexNumber(opts.themeTint);
    }
    if (opts.themeShade !== undefined) {
      attrs["w:themeShade"] = uCharHexNumber(opts.themeShade);
    }
    this.root.push({ _attr: attrs });
  }
}

/**
 * Represents text highlighting in a run.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_HighlightColor">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="black"/>
 *     <xsd:enumeration value="blue"/>
 *     <xsd:enumeration value="cyan"/>
 *     <xsd:enumeration value="green"/>
 *     <xsd:enumeration value="magenta"/>
 *     <xsd:enumeration value="red"/>
 *     <xsd:enumeration value="yellow"/>
 *     <xsd:enumeration value="white"/>
 *     <xsd:enumeration value="darkBlue"/>
 *     <xsd:enumeration value="darkCyan"/>
 *     <xsd:enumeration value="darkGreen"/>
 *     <xsd:enumeration value="darkMagenta"/>
 *     <xsd:enumeration value="darkRed"/>
 *     <xsd:enumeration value="darkYellow"/>
 *     <xsd:enumeration value="darkGray"/>
 *     <xsd:enumeration value="lightGray"/>
 *     <xsd:enumeration value="none"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @example
 * ```typescript
 * new Highlight("yellow"); // Yellow highlight
 * new Highlight("cyan"); // Cyan highlight
 * ```
 *
 * @internal
 */
export class Highlight extends XmlComponent {
  public constructor(color: string) {
    super("w:highlight");
    this.root.push({ _attr: { "w:val": color } });
  }
}

/**
 * Represents text highlighting for complex scripts.
 *
 * Used for highlighting text in complex script languages
 * (e.g., Arabic, Hebrew, Thai).
 *
 * @internal
 */
export class HighlightComplexScript extends XmlComponent {
  public constructor(color: string) {
    super("w:highlightCs");
    this.root.push({ _attr: { "w:val": color } });
  }
}

/**
 * Build character spacing (w:spacing) as IXmlableObject without allocating XmlComponent tree.
 */
export const buildCharacterSpacingObj = (value: number | UniversalMeasure): IXmlableObject =>
  attrObj("w:spacing", { "w:val": signedTwipsMeasureValue(value) });

/**
 * Build text color (w:color) as IXmlableObject without allocating XmlComponent tree.
 */
export const buildColorObj = (colorOrOptions: string | ColorOptions): IXmlableObject => {
  if (typeof colorOrOptions === "string") {
    return attrObj("w:color", { "w:val": hexColorValue(colorOrOptions) });
  }

  const opts = colorOrOptions;
  return attrObj("w:color", {
    "w:val": opts.val === undefined ? undefined : hexColorValue(opts.val),
    "w:themeColor": opts.themeColor,
    "w:themeTint": opts.themeTint === undefined ? undefined : uCharHexNumber(opts.themeTint),
    "w:themeShade": opts.themeShade === undefined ? undefined : uCharHexNumber(opts.themeShade),
  });
};

/**
 * Build highlight (w:highlight) as IXmlableObject without allocating XmlComponent tree.
 */
export const buildHighlightObj = (color: string): IXmlableObject =>
  attrObj("w:highlight", { "w:val": color });

/**
 * Build complex script highlight (w:highlightCs) as IXmlableObject without allocating XmlComponent tree.
 */
export const buildHighlightComplexScriptObj = (color: string): IXmlableObject =>
  attrObj("w:highlightCs", { "w:val": color });
