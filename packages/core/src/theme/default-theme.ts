/**
 * Theme XML generation — pure function with caching.
 *
 * Replaces the DefaultTheme class with a functionally equivalent
 * pure function. Caches by serialized options key.
 *
 * @module
 */
import { buildThemeXml } from "./build-theme-xml";
import type { ThemeOptions } from "./theme-options";

// Pre-computed default theme XML — zero allocation for the common case.
const DEFAULT_XML = buildThemeXml();

// Cache for customized themes, keyed by serialized options.
const customCache = new Map<string, string>();

function themeKey(o: ThemeOptions): string {
  const c = o.colors;
  const f = o.fonts;
  return `n${o.name ?? ""}c${c?.dark1 ?? ""}${c?.light1 ?? ""}${c?.dark2 ?? ""}${c?.light2 ?? ""}${c?.accent1 ?? ""}${c?.accent2 ?? ""}${c?.accent3 ?? ""}${c?.accent4 ?? ""}${c?.accent5 ?? ""}${c?.accent6 ?? ""}${c?.hyperlink ?? ""}${c?.followedHyperlink ?? ""}f${f?.majorFont ?? ""}${f?.minorFont ?? ""}${f?.majorFontAsian ?? ""}${f?.minorFontAsian ?? ""}`;
}

/**
 * Generate theme XML string from options.
 * Returns cached default when no options provided.
 */
export function createThemeXml(options?: ThemeOptions): string {
  if (!options) return DEFAULT_XML;
  const key = themeKey(options);
  const cached = customCache.get(key);
  if (cached) return cached;
  const built = buildThemeXml(options);
  customCache.set(key, built);
  return built;
}
