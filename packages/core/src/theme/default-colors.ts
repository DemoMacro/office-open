/**
 * Default Office 2024 theme colors (Aptos era).
 *
 * Measured from Word 365 (`DocumentTheme.ThemeColorScheme`) on 2026-10-10;
 * the previous values were the Office 2016+ scheme (accent1 4472C4).
 *
 * @module
 */
import type { ColorSchemeOptions } from "./theme-options";

/** Office 2024 default theme colors (hex without #). */
export const DEFAULT_COLORS: Record<Exclude<keyof ColorSchemeOptions, "name">, string> = {
  dark1: "000000",
  light1: "FFFFFF",
  dark2: "0E2841",
  light2: "E8E8E8",
  accent1: "156082",
  accent2: "E97132",
  accent3: "196B24",
  accent4: "0F9ED5",
  accent5: "A02B93",
  accent6: "4EA72E",
  hyperlink: "467886",
  followedHyperlink: "96607D",
};
