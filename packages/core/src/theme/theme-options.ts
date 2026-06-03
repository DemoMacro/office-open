/**
 * Theme options for OOXML documents.
 *
 * @module
 */

/** Color scheme — 12 theme colors (hex without #). */
export interface ColorSchemeOptions {
  readonly dark1?: string;
  readonly light1?: string;
  readonly dark2?: string;
  readonly light2?: string;
  readonly accent1?: string;
  readonly accent2?: string;
  readonly accent3?: string;
  readonly accent4?: string;
  readonly accent5?: string;
  readonly accent6?: string;
  readonly hyperlink?: string;
  readonly followedHyperlink?: string;
}

/** Font scheme — 4 font slots (latin + east-asian for major/minor). */
export interface FontSchemeOptions {
  readonly majorFont?: string;
  readonly minorFont?: string;
  readonly majorFontAsian?: string;
  readonly minorFontAsian?: string;
}

/** Theme customization options. */
export interface ThemeOptions {
  readonly name?: string;
  readonly colors?: ColorSchemeOptions;
  readonly fonts?: FontSchemeOptions;
}
