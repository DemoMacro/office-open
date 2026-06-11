/**
 * Numbering level definitions module for WordprocessingML documents.
 *
 * This module defines the formatting and behavior of individual levels within
 * a numbered or bulleted list hierarchy.
 *
 * Reference: http://officeopenxml.com/WPnumbering-numFmt.php
 *
 * @module
 */
import type { AlignmentType } from "../paragraph/formatting";
import type { LevelParagraphStylePropertiesOptions } from "../paragraph/properties";
import type { RunStylePropertiesOptions } from "../paragraph/run/properties";

/**
 * Number format types for list levels.
 *
 * Defines the various numbering formats available for list levels, including
 * decimal, roman numerals, letters, and various international formats.
 *
 * Reference: http://officeopenxml.com/WPnumbering-numFmt.php
 *
 * @publicApi
 */
export const LevelFormat = {
  /** Decimal numbering (1, 2, 3...). */
  DECIMAL: "decimal",
  /** Uppercase roman numerals (I, II, III...). */
  UPPER_ROMAN: "upperRoman",
  /** Lowercase roman numerals (i, ii, iii...). */
  LOWER_ROMAN: "lowerRoman",
  /** Uppercase letters (A, B, C...). */
  UPPER_LETTER: "upperLetter",
  /** Lowercase letters (a, b, c...). */
  LOWER_LETTER: "lowerLetter",
  /** Ordinal numbers (1st, 2nd, 3rd...). */
  ORDINAL: "ordinal",
  /** Cardinal text (one, two, three...). */
  CARDINAL_TEXT: "cardinalText",
  /** Ordinal text (first, second, third...). */
  ORDINAL_TEXT: "ordinalText",
  /** Hexadecimal numbering. */
  HEX: "hex",
  /** Chicago Manual of Style numbering. */
  CHICAGO: "chicago",
  /** Ideograph digital numbering. */
  IDEOGRAPH__DIGITAL: "ideographDigital",
  /** Japanese counting system. */
  JAPANESE_COUNTING: "japaneseCounting",
  /** Japanese aiueo ordering. */
  AIUEO: "aiueo",
  /** Japanese iroha ordering. */
  IROHA: "iroha",
  /** Full-width decimal numbering. */
  DECIMAL_FULL_WIDTH: "decimalFullWidth",
  /** Half-width decimal numbering. */
  DECIMAL_HALF_WIDTH: "decimalHalfWidth",
  /** Japanese legal numbering. */
  JAPANESE_LEGAL: "japaneseLegal",
  /** Japanese digital ten thousand numbering. */
  JAPANESE_DIGITAL_TEN_THOUSAND: "japaneseDigitalTenThousand",
  /** Decimal numbers enclosed in circles. */
  DECIMAL_ENCLOSED_CIRCLE: "decimalEnclosedCircle",
  /** Full-width decimal numbering variant 2. */
  DECIMAL_FULL_WIDTH2: "decimalFullWidth2",
  /** Full-width aiueo ordering. */
  AIUEO_FULL_WIDTH: "aiueoFullWidth",
  /** Full-width iroha ordering. */
  IROHA_FULL_WIDTH: "irohaFullWidth",
  /** Decimal with leading zeros. */
  DECIMAL_ZERO: "decimalZero",
  /** Bullet points. */
  BULLET: "bullet",
  /** Korean ganada ordering. */
  GANADA: "ganada",
  /** Korean chosung ordering. */
  CHOSUNG: "chosung",
  /** Decimal enclosed with fullstop. */
  DECIMAL_ENCLOSED_FULLSTOP: "decimalEnclosedFullstop",
  /** Decimal enclosed in parentheses. */
  DECIMAL_ENCLOSED_PARENTHESES: "decimalEnclosedParen",
  /** Decimal enclosed in circles (Chinese). */
  DECIMAL_ENCLOSED_CIRCLE_CHINESE: "decimalEnclosedCircleChinese",
  /** Ideograph enclosed in circles. */
  IDEOGRAPH_ENCLOSED_CIRCLE: "ideographEnclosedCircle",
  /** Traditional ideograph numbering. */
  IDEOGRAPH_TRADITIONAL: "ideographTraditional",
  /** Ideograph zodiac numbering. */
  IDEOGRAPH_ZODIAC: "ideographZodiac",
  /** Traditional ideograph zodiac numbering. */
  IDEOGRAPH_ZODIAC_TRADITIONAL: "ideographZodiacTraditional",
  /** Taiwanese counting system. */
  TAIWANESE_COUNTING: "taiwaneseCounting",
  /** Traditional ideograph legal numbering. */
  IDEOGRAPH_LEGAL_TRADITIONAL: "ideographLegalTraditional",
  /** Taiwanese counting thousand system. */
  TAIWANESE_COUNTING_THOUSAND: "taiwaneseCountingThousand",
  /** Taiwanese digital numbering. */
  TAIWANESE_DIGITAL: "taiwaneseDigital",
  /** Chinese counting system. */
  CHINESE_COUNTING: "chineseCounting",
  /** Simplified Chinese legal numbering. */
  CHINESE_LEGAL_SIMPLIFIED: "chineseLegalSimplified",
  /** Chinese counting thousand system. */
  CHINESE_COUNTING_THOUSAND: "chineseCountingThousand",
  /** Korean digital numbering. */
  KOREAN_DIGITAL: "koreanDigital",
  /** Korean counting system. */
  KOREAN_COUNTING: "koreanCounting",
  /** Korean legal numbering. */
  KOREAN_LEGAL: "koreanLegal",
  /** Korean digital numbering variant 2. */
  KOREAN_DIGITAL2: "koreanDigital2",
  /** Vietnamese counting system. */
  VIETNAMESE_COUNTING: "vietnameseCounting",
  /** Russian lowercase numbering. */
  RUSSIAN_LOWER: "russianLower",
  /** Russian uppercase numbering. */
  RUSSIAN_UPPER: "russianUpper",
  /** No numbering. */
  NONE: "none",
  /** Number enclosed in dashes. */
  NUMBER_IN_DASH: "numberInDash",
  /** Hebrew numbering variant 1. */
  HEBREW1: "hebrew1",
  /** Hebrew numbering variant 2. */
  HEBREW2: "hebrew2",
  /** Arabic alpha numbering. */
  ARABIC_ALPHA: "arabicAlpha",
  /** Arabic abjad numbering. */
  ARABIC_ABJAD: "arabicAbjad",
  /** Hindi vowels. */
  HINDI_VOWELS: "hindiVowels",
  /** Hindi consonants. */
  HINDI_CONSONANTS: "hindiConsonants",
  /** Hindi numbers. */
  HINDI_NUMBERS: "hindiNumbers",
  /** Hindi counting system. */
  HINDI_COUNTING: "hindiCounting",
  /** Thai letters. */
  THAI_LETTERS: "thaiLetters",
  /** Thai numbers. */
  THAI_NUMBERS: "thaiNumbers",
  /** Thai counting system. */
  THAI_COUNTING: "thaiCounting",
  /** Thai Baht text. */
  BAHT_TEXT: "bahtText",
  /** Dollar text. */
  DOLLAR_TEXT: "dollarText",
  /** Custom numbering format. */
  CUSTOM: "custom",
} as const;

/**
 * Suffix types for list levels.
 *
 * Defines what follows the numbering text (tab, space, or nothing).
 *
 * @publicApi
 */
export const LevelSuffix = {
  /** No separator after the numbering. */
  NOTHING: "nothing",
  /** Space character after the numbering. */
  SPACE: "space",
  /** Tab character after the numbering. */
  TAB: "tab",
} as const;

/**
 * Options for configuring a numbering level.
 *
 * @property level - Level index (0-8)
 * @property format - Number format type (decimal, roman, letter, etc.)
 * @property text - Level text template with placeholders like %1, %2
 * @property alignment - Text alignment for the numbering
 * @property start - Starting number for this level
 * @property suffix - Character(s) following the numbering (tab, space, nothing)
 * @property isLegalNumberingStyle - Use legal numbering style
 * @property style - Run and paragraph style properties
 */
export interface LevelsOptions {
  /** Level index (0-8). */
  level: number;
  /** Number format type (decimal, roman, letter, bullet, etc.). */
  format?: (typeof LevelFormat)[keyof typeof LevelFormat];
  /** Level text template with placeholders like %1, %2. */
  text?: string;
  /** Text alignment for the numbering. */
  alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
  /** Starting number for this level. */
  start?: number;
  /** Character(s) following the numbering. */
  suffix?: (typeof LevelSuffix)[keyof typeof LevelSuffix];
  /** Use legal numbering style (e.g., 1.1.1). */
  isLegalNumberingStyle?: boolean;
  /** Restart numbering after this level (0-based level index). */
  lvlRestart?: number;
  /** Picture bullet ID reference. */
  lvlPicBulletId?: number;
  /** Template code for the level. */
  templateCode?: string;
  /** Whether this level is tentative. */
  tentative?: boolean;
  /** Legacy spacing/indent settings. */
  legacy?: { space?: number; indent?: number };
  /** Run and paragraph style properties. */
  style?: {
    /** Run style properties for the numbering text. */
    run?: RunStylePropertiesOptions;
    /** Paragraph style properties for the level. */
    paragraph?: LevelParagraphStylePropertiesOptions;
  };
}
