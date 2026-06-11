/**
 * Run fonts module for WordprocessingML documents.
 *
 * This module provides support for specifying fonts for different character sets
 * within a run. Fonts can be specified separately for ASCII, complex script (CS),
 * East Asian, and high ANSI character ranges.
 *
 * Reference: http://officeopenxml.com/WPrun.php
 *
 * @module
 */
import type { ThemeFont } from "@office-open/core";

/**
 * Options for font attributes across different character sets.
 *
 * @property ascii - Font for ASCII characters (0x00-0x7F)
 * @property cs - Font for complex script characters
 * @property eastAsia - Font for East Asian characters
 * @property hAnsi - Font for high ANSI characters (0x80-0xFF)
 * @property hint - Hint for font selection algorithm
 * @property asciiTheme - Theme font for ASCII characters
 * @property hAnsiTheme - Theme font for high ANSI characters
 * @property eastAsiaTheme - Theme font for East Asian characters
 * @property cstheme - Theme font for complex script characters
 */
export interface FontAttributesProperties {
  /** Font for ASCII characters (0x00-0x7F) */
  ascii?: string;
  /** Font for complex script characters */
  cs?: string;
  /** Font for East Asian characters */
  eastAsia?: string;
  /** Font for high ANSI characters (0x80-0xFF) */
  hAnsi?: string;
  /** Hint for font selection algorithm */
  hint?: string;
  /** Theme font for ASCII characters */
  asciiTheme?: (typeof ThemeFont)[keyof typeof ThemeFont];
  /** Theme font for high ANSI characters */
  hAnsiTheme?: (typeof ThemeFont)[keyof typeof ThemeFont];
  /** Theme font for East Asian characters */
  eastAsiaTheme?: (typeof ThemeFont)[keyof typeof ThemeFont];
  /** Theme font for complex script characters */
  cstheme?: (typeof ThemeFont)[keyof typeof ThemeFont];
}
