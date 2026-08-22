/**
 * Run formatting module for WordprocessingML documents.
 *
 * This module provides character-level formatting elements including
 * spacing, color, and highlighting.
 *
 * Reference: http://officeopenxml.com/WPtextFormatting.php
 *
 * @module
 */
import type { ThemeColor } from "@office-open/core";
import type { HexColorOrAuto, UcharHexNumber } from "@office-open/core";

/**
 * Options for theme color configuration.
 *
 * @property val - Explicit hex color value (e.g., "FF0000" or "auto")
 * @property themeColor - Theme color reference
 * @property themeTint - Theme color tint (2-char hex, e.g., "99")
 * @property themeShade - Theme color shade (2-char hex, e.g., "BF")
 */
export interface ColorOptions {
  val?: HexColorOrAuto;
  /** Theme color slot: "dark1"/"light1" text/background, "accent1"–"accent6" theme accents, "hyperlink"/"followedHyperlink". */
  themeColor?: ThemeColor;
  themeTint?: UcharHexNumber;
  themeShade?: UcharHexNumber;
}
