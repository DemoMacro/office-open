/**
 * Font module for WordprocessingML documents.
 *
 * Provides character-set constants for font definitions.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-w_font-1.html
 *
 * @module
 */

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
