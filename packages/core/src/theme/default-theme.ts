/**
 * Theme XML generation — pure function with caching.
 *
 * Fresh output (no options) returns a pre-computed default. Simple customizations
 * (colorScheme / fontScheme only) are cached by serialized key. Complex
 * customizations (formatScheme / objectDefaults / extraColorSchemes) bypass the
 * cache — they are rare (round-trip produces a different theme each time) and
 * require a WriteContext for nested descriptors.
 *
 * @module
 */
import type { WriteContext } from "../descriptor";
import { buildThemeXml } from "./build-theme-xml";
import type { ThemeOptions } from "./theme-options";

// Pre-computed default theme XML — zero allocation for the common case.
const DEFAULT_XML = buildThemeXml();

// Cache for simple customizations, keyed by serialized options.
const customCache = new Map<string, string>();

function isSimpleCustomization(opts: ThemeOptions): boolean {
  return !opts.formatScheme && !opts.objectDefaults && !opts.extraColorSchemes;
}

function simpleKey(o: ThemeOptions): string {
  return `${o.name ?? ""}|${JSON.stringify(o.colorScheme)}|${JSON.stringify(o.fontScheme)}`;
}

/**
 * Generate theme XML string from options.
 * Returns the cached default when no options are provided.
 */
export function createThemeXml(options?: ThemeOptions, ctx?: WriteContext): string {
  if (!options) return DEFAULT_XML;
  if (isSimpleCustomization(options)) {
    const key = simpleKey(options);
    const cached = customCache.get(key);
    if (cached) return cached;
    const built = buildThemeXml(options);
    customCache.set(key, built);
    return built;
  }
  return buildThemeXml(options, ctx);
}
