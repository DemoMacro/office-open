/**
 * Font Table module for WordprocessingML documents.
 *
 * This module provides support for embedding fonts in the document.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-w_fonts.html
 *
 * @module
 */
import type { DataType } from "@office-open/core";

import type { CharacterSet } from "./font";

// <xsd:complexType name="CT_FontsList">
//     <xsd:sequence>
//         <xsd:element name="font" type="CT_Font" minOccurs="0" maxOccurs="unbounded"/>
//     </xsd:sequence>
// </xsd:complexType>

/** Font signature bitfields (Unicode subset + code-page ranges). */
export interface FontSignature {
  usb0: string;
  usb1: string;
  usb2: string;
  usb3: string;
  csb0: string;
  csb1: string;
}

/**
 * Options for embedding a font in the document.
 *
 * `data` is optional: a parsed font table carries metadata-only declarations
 * (name/family/pitch/sig/...) for non-embedded fonts, which have no bytes to
 * embed and must still round-trip into fontTable.xml.
 */
/**
 * An embedded font (fontTable.xml w:font with w:embedRegular parts). Author
 * with name/data only; the round-trip fields (rawOdttf/odttfPath/fontKey/
 * subsetted) carry source-file obfuscation state — do not hand-author.
 */
export interface EmbeddedFontOptions {
  /** Font family name */
  name: string;
  /** Font file data (TTF, OTF): raw bytes, ArrayBuffer, or a base64 data URL. */
  data?: DataType;
  /** Character set/encoding for the font */
  characterSet?: (typeof CharacterSet)[keyof typeof CharacterSet];
  /** IANA character set name (w:charset/`@w:characterSet`, e.g. "ISO-8859-1"). */
  characterSetName?: string;
  /** Font family classification (e.g. "auto", "roman", "swiss") */
  family?: string;
  /**
   * Font is not a TrueType font (w:notTrueType). Word writes the bare element
   * for true; `false` emits w:val="0" explicitly.
   */
  notTrueType?: boolean;
  /** Font pitch (e.g. "fixed", "variable") */
  pitch?: string;
  /** PANOSE-1 classification (10-digit hex string) */
  panose1?: string;
  /** Alternative font name */
  altName?: string;
  /** Font signature (Unicode/code-page bitfields) */
  sig?: FontSignature;
  /**
   * Internal round-trip flag: when set, `data` holds already-obfuscated
   * .odttf bytes that must be copied verbatim instead of re-obfuscated.
   */
  rawOdttf?: boolean;
  /** Internal: original .odttf part path (round-trip); preserves source file name. */
  odttfPath?: string;
  /**
   * Internal round-trip: the obfuscation GUID (`w:fontKey`) read back from
   * fontTable.xml, preserved verbatim so re-generated tables keep the same key.
   * Undefined on authoring — a fresh key is generated for embedded fonts.
   */
  fontKey?: string;
  /**
   * Embedded font is subsetted rather than complete (CT_FontRel `@subsetted`).
   * Round-trip flag preserved from source embedRegular.
   */
  subsetted?: boolean;
  /**
   * Markup-compatibility namespace prefix gating this declaration
   * (`mc:AlternateContent/mc:Choice/@Requires`, e.g. "wpc"). When set, the
   * font is emitted inside an AlternateContent/Choice wrapper.
   */
  requires?: string;
}
