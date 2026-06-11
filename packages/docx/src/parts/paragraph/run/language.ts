/**
 * Language module for WordprocessingML run properties.
 *
 * This module provides support for specifying the language of text content,
 * which affects spell checking, grammar checking, and hyphenation.
 *
 * Reference: http://officeopenxml.com/WPrun.php
 *
 * @module
 */

/**
 * Options for language settings.
 *
 * Specifies the language for different text types within a run.
 * Language codes should follow RFC 1766 (e.g., "en-US", "fr-FR", "ja-JP").
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Language">
 *   <xsd:attribute name="val" type="s:ST_Lang" use="optional"/>
 *   <xsd:attribute name="eastAsia" type="s:ST_Lang" use="optional"/>
 *   <xsd:attribute name="bidi" type="s:ST_Lang" use="optional"/>
 * </xsd:complexType>
 * ```
 *
 * @property value - Language for Latin and complex script text (e.g., "en-US")
 * @property eastAsia - Language for East Asian text (e.g., "ja-JP", "zh-CN")
 * @property bidirectional - Language for bidirectional text (e.g., "ar-SA", "he-IL")
 */
export interface LanguageOptions {
  /** Language for Latin and complex script text (RFC 1766 format, e.g., "en-US") */
  value?: string;
  /** Language for East Asian text (RFC 1766 format, e.g., "ja-JP") */
  eastAsia?: string;
  /** Language for bidirectional text (RFC 1766 format, e.g., "ar-SA") */
  bidirectional?: string;
}
