import type { Guid, LongHexNumber, Panose } from "@office-open/core";
/**
 * Font module for WordprocessingML documents.
 *
 * Provides support for font definitions and embedded fonts.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-w_font-1.html
 *
 * @module
 */
/**
 * Options for a font relationship (embedded font).
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-w_embedRegular-1.html
 *
 * @property id - Relationship to Part
 * @property fontKey - Embedded Font Obfuscation Key (GUID)
 * @property subsetted - Whether the embedded font is subsetted
 */
export interface FontRelationshipOptions {
  /** Relationship to Part */
  id: string;
  /** Embedded Font Obfuscation Key (GUID) */
  fontKey?: Guid;
  /** Whether the embedded font is subsetted */
  subsetted?: boolean;
}

/**
 * Character set constants for font definitions.
 * Maps character set names to their hexadecimal identifiers.
 *
 * @publicApi
 */
export const CharacterSet = {
  ANSI: "00",
  ARABIC: "B2",
  BALTIC: "BA",
  CHINESEBIG5: "88",
  DEFAULT: "01",
  EASTEUROPE: "EE",
  GB_2312: "86",
  GREEK: "A1",
  HANGUL: "81",
  HEBREW: "B1",
  JIS: "80",
  JOHAB: "82",
  MAC: "4D",
  OEM: "FF",
  RUSSIAN: "CC",
  SYMBOL: "02",
  THAI: "DE",
  TURKISH: "A2",
  VIETNAMESE: "A3",
} as const;

/**
 * Options for defining a font in the document.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-w_font-1.html
 *
 * @property name - Font name (required)
 * @property altName - Alternative font name
 * @property panose1 - PANOSE-1 classification
 * @property charset - Character set identifier
 * @property family - Font family
 * @property notTrueType - Whether this is not a TrueType font
 * @property pitch - Font pitch
 * @property sig - Font signature (Unicode and code page ranges)
 * @property embedRegular - Embedded regular font relationship
 * @property embedBold - Embedded bold font relationship
 * @property embedItalic - Embedded italic font relationship
 * @property embedBoldItalic - Embedded bold-italic font relationship
 */
export interface FontEntry {
  /** Font name (required) */
  name: string;
  /** Alternative font name */
  altName?: string;
  /** PANOSE-1 classification */
  panose1?: Panose;
  /** Character set identifier */
  charset?: (typeof CharacterSet)[keyof typeof CharacterSet];
  /** Font family (w:family, ST_FontFamily) */
  family?: "decorative" | "modern" | "roman" | "script" | "swiss" | "auto";
  /** Whether this is not a TrueType font */
  notTrueType?: boolean;
  /** Font pitch (w:pitch, ST_Pitch) */
  pitch?: "fixed" | "variable" | "default";
  /** Font signature (Unicode and code page ranges) */
  sig?: {
    /** Unicode Subset Bitfield 0 */
    usb0: LongHexNumber;
    /** Unicode Subset Bitfield 1 */
    usb1: LongHexNumber;
    /** Unicode Subset Bitfield 2 */
    usb2: LongHexNumber;
    /** Unicode Subset Bitfield 3 */
    usb3: LongHexNumber;
    /** Code Page Bitfield 0 */
    csb0: LongHexNumber;
    /** Code Page Bitfield 1 */
    csb1: LongHexNumber;
  };
  /** Embedded regular font relationship */
  embedRegular?: FontRelationshipOptions;
  /** Embedded bold font relationship */
  embedBold?: FontRelationshipOptions;
  /** Embedded italic font relationship */
  embedItalic?: FontRelationshipOptions;
  /** Embedded bold-italic font relationship */
  embedBoldItalic?: FontRelationshipOptions;
}
