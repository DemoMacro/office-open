/**
 * Theme override and theme manager descriptors.
 *
 * a:themeOverride (CT_BaseStylesOverride) is the root of a themeOverride{n}.xml
 * part — an optional subset of the base styles that a slide/layout/master can
 * reference to deviate from its owning theme. a:themeManager (CT_EmptyElement)
 * is the root of the legacy themeManager part; it carries no content.
 *
 * @module
 */

import { findChild } from "@office-open/xml";

import type { CustomDescriptor, WriteContext } from "../descriptor";
import { parseColorScheme, stringifyColorScheme } from "./color-scheme";
import { parseFontScheme, stringifyFontScheme } from "./font-scheme";
import { parseFormatScheme, stringifyFormatScheme } from "./style-matrix";
import type { ThemeOverrideOptions } from "./theme-options";

const THEME_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

export const themeOverrideDesc: CustomDescriptor<
  ThemeOverrideOptions,
  WriteContext,
  ThemeOverrideOptions
> = {
  kind: "custom",

  stringify(opts, ctx) {
    let content = "";
    if (opts.colorScheme) {
      content += stringifyColorScheme(opts.colorScheme, opts.colorScheme.name ?? "Override");
    }
    if (opts.fontScheme) {
      content += stringifyFontScheme(opts.fontScheme, opts.fontScheme.name ?? "Override");
    }
    if (opts.formatScheme) content += stringifyFormatScheme(opts.formatScheme, ctx);
    return `<a:themeOverride ${THEME_NS}>${content}</a:themeOverride>`;
  },

  parse(el, ctx) {
    const result: Partial<ThemeOverrideOptions> = {};
    const colorScheme = parseColorScheme(findChild(el, "a:clrScheme"));
    if (colorScheme) result.colorScheme = colorScheme;
    const fontScheme = parseFontScheme(findChild(el, "a:fontScheme"));
    if (fontScheme) result.fontScheme = fontScheme;
    const formatScheme = parseFormatScheme(findChild(el, "a:fmtScheme"), ctx);
    if (formatScheme) result.formatScheme = formatScheme;
    return result as ThemeOverrideOptions;
  },
};

/** Options for a themeManager part — CT_EmptyElement has no content or attributes. */
export type ThemeManagerOptions = Record<never, never>;

export const themeManagerDesc: CustomDescriptor<
  ThemeManagerOptions,
  WriteContext,
  ThemeManagerOptions
> = {
  kind: "custom",

  stringify(_opts, _ctx) {
    return `<a:themeManager ${THEME_NS}/>`;
  },

  parse(_el, _ctx) {
    return {};
  },
};
