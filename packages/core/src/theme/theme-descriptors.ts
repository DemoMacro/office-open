/**
 * Theme descriptor — bidirectional mapping between ThemeOptions and theme XML.
 *
 * stringify delegates to buildThemeXml (structured driver). parse reads the full
 * CT_BaseStyles surface: colorScheme, fontScheme, formatScheme, objectDefaults,
 * and extraClrSchemeLst — so every part round-trips.
 *
 * @module
 */

import { attr, findChild } from "@office-open/xml";
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

    if (el.name === "a:officeStyleSheet") result.rootElement = "officeStyleSheet";
    if (attr(el, "thm15:id") !== undefined) result.themeExtensionId = attr(el, "thm15:id");
    const elementsContainer = findChild(el, "a:themeElements") ? "themeElements" : "baseStyles";
    const themeElements = findChild(el, `a:${elementsContainer}`);
    if (themeElements) result.elementsContainer = elementsContainer;
    if (themeElements) {
      const colorScheme = parseColorScheme(findChild(themeElements, "a:clrScheme"), ctx);
      if (colorScheme) result.colorScheme = colorScheme;
      const fontScheme = parseFontScheme(findChild(themeElements, "a:fontScheme"));
      if (fontScheme) result.fontScheme = fontScheme;
      const formatScheme = parseFormatScheme(findChild(themeElements, "a:fmtScheme"), ctx);
      if (formatScheme) result.formatScheme = formatScheme;
    }

    const objectDefaults = parseObjectDefaults(findChild(el, "a:objectDefaults"), ctx);
    if (objectDefaults) result.objectDefaults = objectDefaults;
    result.objectDefaultsPresent = findChild(el, "a:objectDefaults") !== undefined;

    const extraColorSchemeList = findChild(el, "a:extraClrSchemeLst");
    const extraColorSchemes = parseExtraColorSchemes(extraColorSchemeList, ctx);
    if (extraColorSchemes) result.extraColorSchemes = extraColorSchemes;
    result.includeExtraColorSchemes = extraColorSchemeList !== undefined;

    const customColors = parseCustomColors(findChild(el, "a:custClrLst"), ctx);
    if (customColors) result.customColors = customColors;

    const themeFamily = findChild(el, "a:extLst")
      ?.elements?.find(
        (child) =>
          child.type === "element" &&
          child.name === "a:ext" &&
          attr(child, "uri") === "{05A4C25C-085E-4340-85A3-A5531E510DB2}",
      )
      ?.elements?.find((child) => child.type === "element" && child.name === "thm15:themeFamily");
    if (themeFamily) {
      const name = attr(themeFamily, "name");
      const id = attr(themeFamily, "id");
      const vid = attr(themeFamily, "vid");
      if (name !== undefined && id !== undefined && vid !== undefined) {
        result.themeFamily = { name, id, vid };
      }
    }

    return result as ThemeOptions;
  },
};
