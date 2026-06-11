/**
 * Font Table module for WordprocessingML documents.
 *
 * This module provides support for embedding fonts in the document.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-w_fonts.html
 *
 * @module
 */
import { element } from "@office-open/xml";

import { createRegularFont } from "./create-regular-font";
import type { CharacterSet } from "./font";
import type { EmbeddedFontOptionsWithKey } from "./font-wrapper";

// <xsd:complexType name="CT_FontsList">
//     <xsd:sequence>
//         <xsd:element name="font" type="CT_Font" minOccurs="0" maxOccurs="unbounded"/>
//     </xsd:sequence>
// </xsd:complexType>

/**
 * Options for embedding a font in the document.
 */
export interface EmbeddedFontOptions {
  /** Font family name */
  name: string;
  /** Font file data (TTF, OTF, etc.) */
  data: Buffer;
  /** Character set/encoding for the font */
  characterSet?: (typeof CharacterSet)[keyof typeof CharacterSet];
}

/**
 * Creates a font table element containing embedded fonts.
 *
 * The font table allows custom fonts to be embedded in the document
 * so they display correctly even if not installed on the viewer's system.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-w_fonts.html
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_FontsList">
 *   <xsd:sequence>
 *     <xsd:element name="font" type="CT_Font" minOccurs="0" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createFontTable = (fonts: EmbeddedFontOptionsWithKey[]): string =>
  // https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_Font_topic_ID0ERNCU.html
  // http://www.datypic.com/sc/ooxml/e-w_fonts.html
  element(
    "w:fonts",
    {
      "mc:Ignorable": "w14 w15 w16se w16cid w16 w16cex w16sdtdh",
      "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
      "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
      "xmlns:w15": "http://schemas.microsoft.com/office/word/2012/wordml",
      "xmlns:w16": "http://schemas.microsoft.com/office/word/2018/wordml",
      "xmlns:w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
      "xmlns:w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
      "xmlns:w16sdtdh": "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
      "xmlns:w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    },
    fonts.map((font, i) =>
      createRegularFont({
        characterSet: font.characterSet,
        fontKey: font.fontKey,
        index: i + 1,
        name: font.name,
      }),
    ),
  );
