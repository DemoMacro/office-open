/**
 * Theme options for OOXML documents.
 *
 * @module
 */

/** Color scheme — 12 theme colors (hex without #). */
export interface ColorSchemeOptions {
  dark1?: string;
  light1?: string;
  dark2?: string;
  light2?: string;
  accent1?: string;
  accent2?: string;
  accent3?: string;
  accent4?: string;
  accent5?: string;
  accent6?: string;
  hyperlink?: string;
  followedHyperlink?: string;
}

/** Font scheme — 4 font slots (latin + east-asian for major/minor). */
export interface FontSchemeOptions {
  majorFont?: string;
  minorFont?: string;
  majorFontAsian?: string;
  minorFontAsian?: string;
}

/** Theme customization options. */
export interface ThemeOptions {
  name?: string;
  colors?: ColorSchemeOptions;
  fonts?: FontSchemeOptions;
}
