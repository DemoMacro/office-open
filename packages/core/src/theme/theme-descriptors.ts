/**
 * Theme descriptors — bidirectional mapping between ThemeOptions and theme XML.
 *
 * @module
 */

import { findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type { CustomDescriptor, ReadContext, WriteContext } from "../descriptor";
import { buildThemeXml } from "./build-theme-xml";
import type { ColorSchemeOptions, FontSchemeOptions, ThemeOptions } from "./theme-options";

/** Mutable version for building read results. */
type Mutable<T> = { -readonly [P in keyof T]?: T[P] };

/** Color scheme XML tag → ThemeOptions.colors key mapping. */
const COLOR_TAGS: ReadonlyArray<{
  readonly tag: string;
  readonly key: keyof ColorSchemeOptions & string;
}> = [
  { tag: "a:dk1", key: "dark1" },
  { tag: "a:lt1", key: "light1" },
  { tag: "a:dk2", key: "dark2" },
  { tag: "a:lt2", key: "light2" },
  { tag: "a:accent1", key: "accent1" },
  { tag: "a:accent2", key: "accent2" },
  { tag: "a:accent3", key: "accent3" },
  { tag: "a:accent4", key: "accent4" },
  { tag: "a:accent5", key: "accent5" },
  { tag: "a:accent6", key: "accent6" },
  { tag: "a:hlink", key: "hyperlink" },
  { tag: "a:folHlink", key: "followedHyperlink" },
];

/** Read the hex color value from a color element (a:dk1/a:lt1 may be sysClr or srgbClr). */
function readColorValue(el: XmlElement | undefined): string | undefined {
  if (!el) return undefined;
  const srgb = findChild(el, "a:srgbClr");
  if (srgb) return String(srgb.attributes?.["val"] ?? "");
  const sysClr = findChild(el, "a:sysClr");
  if (sysClr) return String(sysClr.attributes?.["lastClr"] ?? "");
  return undefined;
}

export const themeDesc: CustomDescriptor<ThemeOptions> = {
  kind: "custom",

  stringify(opts: ThemeOptions, _ctx: WriteContext): string | undefined {
    return buildThemeXml(opts);
  },

  parse(el: XmlElement, _ctx: ReadContext) {
    const result: Mutable<ThemeOptions> = {};

    // Theme name
    const name = el.attributes?.["name"];
    if (name) result.name = String(name);

    // Navigate to themeElements > clrScheme
    const themeElements = findChild(el, "a:themeElements");
    const clrScheme = themeElements ? findChild(themeElements, "a:clrScheme") : undefined;

    if (clrScheme) {
      const colors: Mutable<ColorSchemeOptions> = {};
      for (const { tag, key } of COLOR_TAGS) {
        const child = findChild(clrScheme, tag);
        const value = readColorValue(child);
        if (value) colors[key] = value;
      }
      if (Object.keys(colors).length > 0) {
        result.colors = colors as ColorSchemeOptions;
      }
    }

    // Navigate to themeElements > fontScheme
    const fontScheme = themeElements ? findChild(themeElements, "a:fontScheme") : undefined;
    if (fontScheme) {
      const majorFontEl = findChild(fontScheme, "a:majorFont");
      const minorFontEl = findChild(fontScheme, "a:minorFont");
      const majorLatin = majorFontEl ? findChild(majorFontEl, "a:latin") : undefined;
      const minorLatin = minorFontEl ? findChild(minorFontEl, "a:latin") : undefined;

      const majorTypeface = majorLatin?.attributes?.["typeface"];
      const minorTypeface = minorLatin?.attributes?.["typeface"];

      if (majorTypeface || minorTypeface) {
        const fonts: Mutable<FontSchemeOptions> = {};
        if (majorTypeface) fonts.majorFont = String(majorTypeface);
        if (minorTypeface) fonts.minorFont = String(minorTypeface);
        result.fonts = fonts as FontSchemeOptions;
      }
    }

    return result as ThemeOptions;
  },
};
