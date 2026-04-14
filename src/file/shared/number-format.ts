/**
 * Number format module for WordprocessingML documents.
 *
 * This module provides number format constants for page numbers,
 * list numbering, and other numbered elements.
 *
 * Reference: http://officeopenxml.com/WPnumbering-numFmt.php
 *
 * @module
 */

/**
 * Number format types for page numbers and list numbering.
 *
 * Provides international number formats including decimal, Roman numerals,
 * alphabetic, and various Asian numbering systems.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_NumberFormat">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="decimal"/>
 *     <xsd:enumeration value="upperRoman"/>
 *     <xsd:enumeration value="lowerRoman"/>
 *     <xsd:enumeration value="upperLetter"/>
 *     <xsd:enumeration value="lowerLetter"/>
 *     <!-- ... many more formats ... -->
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @example
 * ```typescript
 * // Arabic numerals (1, 2, 3)
 * NumberFormat.DECIMAL;
 *
 * // Roman numerals (I, II, III)
 * NumberFormat.UPPER_ROMAN;
 *
 * // Letters (a, b, c)
 * NumberFormat.LOWER_LETTER;
 * ```
 *
 * @publicApi
 */
export const NumberFormat = {
    AIUEO: "aiueo",
    AIUEO_FULL_WIDTH: "aiueoFullWidth",
    ARABIC_ABJAD: "arabicAbjad",
    ARABIC_ALPHA: "arabicAlpha",
    BAHT_TEXT: "bahtText",
    BULLET: "bullet",
    CARDINAL_TEXT: "cardinalText",
    CHICAGO: "chicago",
    CHINESE_COUNTING: "chineseCounting",
    CHINESE_COUNTING_TEN_THOUSAND: "chineseCountingThousand",
    CHINESE_LEGAL_SIMPLIFIED: "chineseLegalSimplified",
    CHOSUNG: "chosung",
    DECIMAL: "decimal",
    DECIMAL_ENCLOSED_CIRCLE: "decimalEnclosedCircle",
    DECIMAL_ENCLOSED_CIRCLE_CHINESE: "decimalEnclosedCircleChinese",
    DECIMAL_ENCLOSED_FULL_STOP: "decimalEnclosedFullstop",
    DECIMAL_ENCLOSED_PAREN: "decimalEnclosedParen",
    DECIMAL_FULL_WIDTH: "decimalFullWidth",
    DECIMAL_FULL_WIDTH_2: "decimalFullWidth2",
    DECIMAL_HALF_WIDTH: "decimalHalfWidth",
    DECIMAL_ZERO: "decimalZero",
    DOLLAR_TEXT: "dollarText",
    GANADA: "ganada",
    HEBREW_1: "hebrew1",
    HEBREW_2: "hebrew2",
    HEX: "hex",
    HINDI_CONSONANTS: "hindiConsonants",
    HINDI_COUNTING: "hindiCounting",
    HINDI_NUMBERS: "hindiNumbers",
    HINDI_VOWELS: "hindiVowels",
    IDEOGRAPH_DIGITAL: "ideographDigital",
    IDEOGRAPH_ENCLOSED_CIRCLE: "ideographEnclosedCircle",
    IDEOGRAPH_LEGAL_TRADITIONAL: "ideographLegalTraditional",
    IDEOGRAPH_TRADITIONAL: "ideographTraditional",
    IDEOGRAPH_ZODIAC: "ideographZodiac",
    IDEOGRAPH_ZODIAC_TRADITIONAL: "ideographZodiacTraditional",
    IROHA: "iroha",
    IROHA_FULL_WIDTH: "irohaFullWidth",
    JAPANESE_COUNTING: "japaneseCounting",
    JAPANESE_DIGITAL_TEN_THOUSAND: "japaneseDigitalTenThousand",
    JAPANESE_LEGAL: "japaneseLegal",
    KOREAN_COUNTING: "koreanCounting",
    KOREAN_DIGITAL: "koreanDigital",
    KOREAN_DIGITAL_2: "koreanDigital2",
    KOREAN_LEGAL: "koreanLegal",
    LOWER_LETTER: "lowerLetter",
    LOWER_ROMAN: "lowerRoman",
    NONE: "none",
    NUMBER_IN_DASH: "numberInDash",
    ORDINAL: "ordinal",
    ORDINAL_TEXT: "ordinalText",
    RUSSIAN_LOWER: "russianLower",
    RUSSIAN_UPPER: "russianUpper",
    TAIWANESE_COUNTING: "taiwaneseCounting",
    TAIWANESE_COUNTING_THOUSAND: "taiwaneseCountingThousand",
    TAIWANESE_DIGITAL: "taiwaneseDigital",
    THAI_COUNTING: "thaiCounting",
    THAI_LETTERS: "thaiLetters",
    THAI_NUMBERS: "thaiNumbers",
    UPPER_LETTER: "upperLetter",
    UPPER_ROMAN: "upperRoman",
    VIETNAMESE_COUNTING: "vietnameseCounting",
    //   <xsd:enumeration value="custom"/>
} as const;
