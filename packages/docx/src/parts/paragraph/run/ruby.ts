/**
 * Ruby annotation module for WordprocessingML documents.
 *
 * Ruby annotations are small text rendered above or below base text,
 * commonly used for pronunciation guides in East Asian languages (furigana
 * in Japanese, pinyin in Chinese, etc.).
 *
 * Used via the JSON API: `{ ruby: { text: "かな", base: "漢字" } }` as a
 * paragraph or run child.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_Ruby
 *
 * @module
 */

/**
 * Ruby alignment options.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_RubyAlign">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="center"/>
 *     <xsd:enumeration value="distributeLetter"/>
 *     <xsd:enumeration value="distributeSpace"/>
 *     <xsd:enumeration value="left"/>
 *     <xsd:enumeration value="right"/>
 *     <xsd:enumeration value="rightVertical"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 */
export const RubyAlign = {
  /** Center alignment */
  CENTER: "center",
  /** Distribute letters evenly */
  DISTRIBUTE_LETTER: "distributeLetter",
  /** Distribute space evenly */
  DISTRIBUTE_SPACE: "distributeSpace",
  /** Left alignment */
  LEFT: "left",
  /** Right alignment */
  RIGHT: "right",
  /** Right vertical alignment */
  RIGHT_VERTICAL: "rightVertical",
} as const;

/**
 * Options for a Ruby annotation.
 *
 * Used as `{ ruby: RubyOptions }` in paragraph or run children.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Ruby">
 *   <xsd:sequence>
 *     <xsd:element name="rubyPr" type="CT_RubyPr"/>
 *     <xsd:element name="rt" type="CT_RubyContent"/>
 *     <xsd:element name="rubyBase" type="CT_RubyContent"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Japanese furigana
 * { ruby: {
 *   text: "かな",
 *   base: "漢字",
 *   alignment: "center",
 *   fontSize: 20,
 *   languageId: "ja-JP",
 * }}
 * ```
 */
export interface RubyOptions {
  /** Ruby annotation text (e.g., furigana, pinyin) */
  text: string;
  /** Base text being annotated */
  base: string;
  /** Ruby alignment (defaults to "center"). */
  alignment?: (typeof RubyAlign)[keyof typeof RubyAlign];
  /**
   * Font size for the ruby annotation text in points (e.g., 10 = 10pt).
   * Defaults to half the base text size if not specified.
   */
  fontSize?: number;
  /**
   * Vertical offset for the ruby annotation in points.
   * How far the annotation is raised above (or below) the base text.
   */
  raise?: number;
  /**
   * Font size for the base text in points.
   * Used to calculate the ruby annotation positioning.
   */
  baseFontSize?: number;
  /** Language identifier for the ruby annotation (e.g., "ja-JP"). */
  languageId?: string;
  /** Whether the ruby annotation is dirty (needs recalculation). */
  dirty?: boolean;
}
