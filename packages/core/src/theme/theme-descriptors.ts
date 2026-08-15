/**
 * Theme descriptor — bidirectional mapping between ThemeOptions and theme XML.
 *
 * stringify delegates to buildThemeXml (structured driver). parse reads the full
 * CT_BaseStyles surface: colorScheme, fontScheme, formatScheme, objectDefaults,
 * and extraClrSchemeLst — so every part round-trips.
 *
 * @module
 */

import { findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type { CustomDescriptor, ReadContext, WriteContext } from "../descriptor";
import { buildThemeXml } from "./build-theme-xml";
import { parseColorScheme } from "./color-scheme";
import { parseCustomColors } from "./custom-color";
import { parseExtraColorSchemes } from "./extra-color-scheme";
import { parseFontScheme } from "./font-scheme";
import { parseObjectDefaults } from "./object-defaults";
import { parseFormatScheme } from "./style-matrix";
import type { ThemeOptions } from "./theme-options";

export const themeDesc: CustomDescriptor<ThemeOptions, WriteContext, ThemeOptions> = {
  kind: "custom",

  stringify(opts: ThemeOptions, ctx: WriteContext): string | undefined {
    return buildThemeXml(opts, ctx);
  },

  parse(el: XmlElement, ctx: ReadContext): ThemeOptions {
    const result: Partial<ThemeOptions> = {};

    const name = el.attributes?.["name"];
    if (name) result.name = String(name);

    const themeElements = findChild(el, "a:themeElements");
    if (themeElements) {
      const colorScheme = parseColorScheme(findChild(themeElements, "a:clrScheme"));
      if (colorScheme) result.colorScheme = colorScheme;
      const fontScheme = parseFontScheme(findChild(themeElements, "a:fontScheme"));
      if (fontScheme) result.fontScheme = fontScheme;
      const formatScheme = parseFormatScheme(findChild(themeElements, "a:fmtScheme"), ctx);
      if (formatScheme) result.formatScheme = formatScheme;
    }

    const objectDefaults = parseObjectDefaults(findChild(el, "a:objectDefaults"), ctx);
    if (objectDefaults) result.objectDefaults = objectDefaults;

    const extraColorSchemes = parseExtraColorSchemes(findChild(el, "a:extraClrSchemeLst"));
    if (extraColorSchemes) result.extraColorSchemes = extraColorSchemes;

    const customColors = parseCustomColors(findChild(el, "a:custClrLst"), ctx);
    if (customColors) result.customColors = customColors;

    return result as ThemeOptions;
  },
};
