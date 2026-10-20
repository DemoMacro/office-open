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
import type { UcharHexNumber } from "@office-open/core";

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

/** One embedded font face (w:embedRegular, w:embedBold, …). */
export interface EmbeddedFontPartOptions {
  /** Font file data (TTF, OTF): raw bytes, ArrayBuffer, or a base64 data URL. */
  data?: DataType;
  /**
   * Round-trip only: `data` holds already-obfuscated .odttf bytes that must
   * be copied verbatim instead of re-obfuscated.
   */
  rawOdttf?: boolean;
  /** Round-trip only: original .odttf part path; preserves the source name. */
  odttfPath?: string;
  /** Round-trip only: source fontTable relationship id. */
  embedRid?: string;
  /**
   * Round-trip only: obfuscation GUID (`w:fontKey`) preserved verbatim.
   * Undefined on authoring — a fresh key is generated for embedded fonts.
   */
  fontKey?: string;
  /** Embedded font is subsetted rather than complete (CT_FontRel `@subsetted`). */
  subsetted?: boolean;
}

/**
 * A font declaration. Author with `name` and up to four embedded faces; each
 * face's round-trip fields carry source-file obfuscation state — do not
 * hand-author.
 */
export interface EmbeddedFontOptions {
  /** Font family name */
  name: string;
  /** Source `w:font/@w:name` when absent or an empty string; round-trip only. */
  nameRaw?: string;
  /** Character set/encoding for the font */
  characterSet?: UcharHexNumber;
  /** Source `w:charset/@w:val` when it is not 2-char hex; round-trip only. */
  characterSetRaw?: string;
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
  /**
   * Font signature (Unicode/code-page bitfields); null preserves a source
   * declaration that omitted w:sig.
   */
  sig?: FontSignature | null;
  /** Embedded regular face (w:embedRegular). */
  embedRegular?: EmbeddedFontPartOptions;
  /** Embedded bold face (w:embedBold). */
  embedBold?: EmbeddedFontPartOptions;
  /** Embedded italic face (w:embedItalic). */
  embedItalic?: EmbeddedFontPartOptions;
  /** Embedded bold-italic face (w:embedBoldItalic). */
  embedBoldItalic?: EmbeddedFontPartOptions;
  /**
   * Markup-compatibility namespace prefix gating this declaration
   * (`mc:AlternateContent/mc:Choice/@Requires`, e.g. "wpc"). When set, the
   * font is emitted inside an AlternateContent/Choice wrapper.
   */
  requires?: string;
}

/** Embedded face element names in CT_Font order. */
export const EMBEDDED_FONT_PART_NAMES = [
  "embedRegular",
  "embedBold",
  "embedItalic",
  "embedBoldItalic",
] as const;

export type EmbeddedFontPartName = (typeof EMBEDDED_FONT_PART_NAMES)[number];
