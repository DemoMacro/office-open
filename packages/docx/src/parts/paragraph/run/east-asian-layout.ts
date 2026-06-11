/**
 * East Asian layout module for WordprocessingML run properties.
 *
 * Specifies East Asian typography settings for a run, including
 * character combination, vertical text, and compression.
 *
 * Reference: ISO/IEC 29500-4, CT_EastAsianLayout
 *
 * @module
 */

/**
 * Bracket types for combined characters in East Asian layout.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_CombineBrackets">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="none"/>
 *     <xsd:enumeration value="round"/>
 *     <xsd:enumeration value="square"/>
 *     <xsd:enumeration value="angle"/>
 *     <xsd:enumeration value="curly"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @publicApi
 */
export const CombineBracketsType = {
  /** No brackets */
  NONE: "none",
  /** Round brackets (parentheses) */
  ROUND: "round",
  /** Square brackets */
  SQUARE: "square",
  /** Angle brackets */
  ANGLE: "angle",
  /** Curly brackets (braces) */
  CURLY: "curly",
} as const;

/**
 * Options for East Asian layout properties.
 */
export interface EastAsianLayoutOptions {
  /** Character combination ID (referencing a combined character entry) */
  id?: number;
  /** Whether to combine characters */
  combine?: boolean;
  /** Bracket type for combined characters */
  combineBrackets?: (typeof CombineBracketsType)[keyof typeof CombineBracketsType];
  /** Whether to render text vertically */
  vert?: boolean;
  /** Whether to compress characters in vertical text */
  vertCompress?: boolean;
}
