/**
 * Theme XML generation — pure function.
 *
 * Fresh output (no options) returns a pre-computed default; anything else
 * rebuilds via buildThemeXml (simple customizations don't need a WriteContext).
 *
 * @module
 */
import type { WriteContext } from "../descriptor";
import { buildThemeXml } from "./build-theme-xml";
import type { ThemeOptions } from "./theme-options";

// Pre-computed default theme XML — zero allocation for the common case.
const DEFAULT_XML = buildThemeXml();

/**
 * Generate theme XML string from options.
 * Returns the pre-computed default when no options are provided.
 */
export function createThemeXml(options?: ThemeOptions, ctx?: WriteContext): string {
  if (!options) return DEFAULT_XML;
  return buildThemeXml(options, ctx);
}
