/**
 * Run fonts module for WordprocessingML documents.
 *
 * This module provides support for specifying fonts for different character sets
 * within a run. Fonts can be specified separately for ASCII, complex script (CS),
 * East Asian, and high ANSI character ranges.
 *
 * Reference: http://officeopenxml.com/WPrun.php
 *
 * @module
 */
import { BuilderElement, attrObj } from "@file/xml-components";
import type { IXmlableObject, XmlComponent } from "@file/xml-components";
import type { ThemeFont } from "@util/values";

/**
 * Options for font attributes across different character sets.
 *
 * @property ascii - Font for ASCII characters (0x00-0x7F)
 * @property cs - Font for complex script characters
 * @property eastAsia - Font for East Asian characters
 * @property hAnsi - Font for high ANSI characters (0x80-0xFF)
 * @property hint - Hint for font selection algorithm
 * @property asciiTheme - Theme font for ASCII characters
 * @property hAnsiTheme - Theme font for high ANSI characters
 * @property eastAsiaTheme - Theme font for East Asian characters
 * @property cstheme - Theme font for complex script characters
 */
export interface FontAttributesProperties {
  /** Font for ASCII characters (0x00-0x7F) */
  readonly ascii?: string;
  /** Font for complex script characters */
  readonly cs?: string;
  /** Font for East Asian characters */
  readonly eastAsia?: string;
  /** Font for high ANSI characters (0x80-0xFF) */
  readonly hAnsi?: string;
  /** Hint for font selection algorithm */
  readonly hint?: string;
  /** Theme font for ASCII characters */
  readonly asciiTheme?: (typeof ThemeFont)[keyof typeof ThemeFont];
  /** Theme font for high ANSI characters */
  readonly hAnsiTheme?: (typeof ThemeFont)[keyof typeof ThemeFont];
  /** Theme font for East Asian characters */
  readonly eastAsiaTheme?: (typeof ThemeFont)[keyof typeof ThemeFont];
  /** Theme font for complex script characters */
  readonly cstheme?: (typeof ThemeFont)[keyof typeof ThemeFont];
}

/**
 * Creates font settings for a run in a WordprocessingML document.
 *
 * The rFonts element specifies which fonts should be used for different character
 * ranges in a run. This allows documents to use different fonts for ASCII, complex
 * script, East Asian, and high ANSI characters.
 *
 * Reference: http://officeopenxml.com/WPrun.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Fonts">
 *   <xsd:attribute name="hint" type="ST_Hint"/>
 *   <xsd:attribute name="ascii" type="s:ST_String"/>
 *   <xsd:attribute name="hAnsi" type="s:ST_String"/>
 *   <xsd:attribute name="eastAsia" type="s:ST_String"/>
 *   <xsd:attribute name="cs" type="s:ST_String"/>
 *   <xsd:attribute name="asciiTheme" type="ST_Theme"/>
 *   <xsd:attribute name="hAnsiTheme" type="ST_Theme"/>
 *   <xsd:attribute name="eastAsiaTheme" type="ST_Theme"/>
 *   <xsd:attribute name="cstheme" type="ST_Theme"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Use same font for all character sets
 * createRunFonts("Arial");
 *
 * // Specify different fonts for different character sets
 * createRunFonts({
 *   ascii: "Arial",
 *   eastAsia: "MS Mincho",
 *   cs: "Arial",
 *   hAnsi: "Arial",
 * });
 * ```
 */
export const createRunFonts = (
  nameOrAttrs: string | FontAttributesProperties,
  hint?: string,
): XmlComponent => {
  if (typeof nameOrAttrs === "string") {
    const name = nameOrAttrs;
    return new BuilderElement<FontAttributesProperties>({
      attributes: {
        ascii: { key: "w:ascii", value: name },
        cs: { key: "w:cs", value: name },
        eastAsia: { key: "w:eastAsia", value: name },
        hAnsi: { key: "w:hAnsi", value: name },
        hint: { key: "w:hint", value: hint },
      },
      name: "w:rFonts",
    });
  }

  const attrs = nameOrAttrs;
  return new BuilderElement<FontAttributesProperties>({
    attributes: {
      ascii: { key: "w:ascii", value: attrs.ascii },
      asciiTheme: { key: "w:asciiTheme", value: attrs.asciiTheme },
      cs: { key: "w:cs", value: attrs.cs },
      cstheme: { key: "w:cstheme", value: attrs.cstheme },
      eastAsia: { key: "w:eastAsia", value: attrs.eastAsia },
      eastAsiaTheme: { key: "w:eastAsiaTheme", value: attrs.eastAsiaTheme },
      hAnsi: { key: "w:hAnsi", value: attrs.hAnsi },
      hAnsiTheme: { key: "w:hAnsiTheme", value: attrs.hAnsiTheme },
      hint: { key: "w:hint", value: attrs.hint },
    },
    name: "w:rFonts",
  });
};

/**
 * Build run fonts (w:rFonts) as IXmlableObject without allocating XmlComponent tree.
 */
export const buildRunFontsObj = (
  nameOrAttrs: string | FontAttributesProperties,
  hint?: string,
): IXmlableObject => {
  if (typeof nameOrAttrs === "string") {
    const name = nameOrAttrs;
    return attrObj("w:rFonts", {
      "w:ascii": name,
      "w:cs": name,
      "w:eastAsia": name,
      "w:hAnsi": name,
      "w:hint": hint,
    });
  }

  const attrs = nameOrAttrs;
  return attrObj("w:rFonts", {
    "w:ascii": attrs.ascii,
    "w:asciiTheme": attrs.asciiTheme,
    "w:cs": attrs.cs,
    "w:cstheme": attrs.cstheme,
    "w:eastAsia": attrs.eastAsia,
    "w:eastAsiaTheme": attrs.eastAsiaTheme,
    "w:hAnsi": attrs.hAnsi,
    "w:hAnsiTheme": attrs.hAnsiTheme,
    "w:hint": attrs.hint,
  });
};
