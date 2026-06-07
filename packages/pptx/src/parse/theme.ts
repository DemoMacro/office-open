import { attr, colorAttr, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { ColorSchemeOptions, FontSchemeOptions, ThemeOptions } from "../file/theme";

// Mapping: XML element name → ColorSchemeOptions key
const COLOR_MAP: Record<string, keyof ColorSchemeOptions> = {
  "a:dk1": "dark1",
  "a:lt1": "light1",
  "a:dk2": "dark2",
  "a:lt2": "light2",
  "a:accent1": "accent1",
  "a:accent2": "accent2",
  "a:accent3": "accent3",
  "a:accent4": "accent4",
  "a:accent5": "accent5",
  "a:accent6": "accent6",
  "a:hlink": "hyperlink",
  "a:folHlink": "followedHyperlink",
};

function extractColor(el: Element): string | undefined {
  const srgb = findChild(el, "a:srgbClr");
  if (srgb) return colorAttr(srgb, "val") ?? undefined;
  const sysClr = findChild(el, "a:sysClr");
  if (sysClr) return attr(sysClr, "lastClr") ?? undefined;
  return undefined;
}

export function parseColorScheme(el: Element): ColorSchemeOptions {
  const colors: Record<string, string> = {};
  for (const child of el.elements ?? []) {
    if (!child.name) continue;
    const key = COLOR_MAP[child.name];
    if (!key) continue;
    const val = extractColor(child);
    if (val) colors[key] = val;
  }
  return colors as ColorSchemeOptions;
}

export function parseFontScheme(el: Element): FontSchemeOptions {
  const fonts: Record<string, string> = {};
  const major = findChild(el, "a:majorFont");
  if (major) {
    const latin = findChild(major, "a:latin");
    if (latin) fonts.majorFont = attr(latin, "typeface") ?? "";
    const ea = findChild(major, "a:ea");
    if (ea) fonts.majorFontAsian = attr(ea, "typeface") ?? "";
  }
  const minor = findChild(el, "a:minorFont");
  if (minor) {
    const latin = findChild(minor, "a:latin");
    if (latin) fonts.minorFont = attr(latin, "typeface") ?? "";
    const ea = findChild(minor, "a:ea");
    if (ea) fonts.minorFontAsian = attr(ea, "typeface") ?? "";
  }
  return fonts as FontSchemeOptions;
}

export function parseTheme(el: Element): ThemeOptions {
  const opts: Record<string, unknown> = {};
  const name = attr(el, "name");
  if (name) opts.name = name;

  const themeElements = findChild(el, "a:themeElements");
  if (themeElements) {
    const clrScheme = findChild(themeElements, "a:clrScheme");
    if (clrScheme) opts.colors = parseColorScheme(clrScheme);

    const fontScheme = findChild(themeElements, "a:fontScheme");
    if (fontScheme) opts.fonts = parseFontScheme(fontScheme);
  }

  return opts as ThemeOptions;
}
