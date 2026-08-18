/**
 * Extra color scheme (a:extraClrScheme / CT_ExtraColorScheme) stringify + parse.
 *
 * Each entry pairs a color scheme with an optional color mapping. The list
 * wrapper (a:extraClrSchemeLst / CT_ColorSchemeList) holds zero or more entries.
 *
 * @module
 */
import { findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type { ReadContext } from "../descriptor";
import { parseColorMapping, stringifyColorMapping } from "./color-mapping";
import { parseColorScheme, stringifyColorScheme } from "./color-scheme";
import type { ExtraColorSchemeOptions } from "./theme-options";

/** Serialize a:extraClrSchemeLst. Returns an empty element when no schemes set. */
export function stringifyExtraColorSchemes(
  opts: ExtraColorSchemeOptions[] | undefined,
  fallbackName: string,
): string {
  if (!opts || opts.length === 0) return "<a:extraClrSchemeLst/>";
  const items = opts
    .map((scheme) => {
      const clrScheme = stringifyColorScheme(
        scheme.colorScheme,
        scheme.colorScheme.name ?? fallbackName,
      );
      const clrMap = scheme.colorMapping ? stringifyColorMapping(scheme.colorMapping) : "";
      return `<a:extraClrScheme>${clrScheme}${clrMap}</a:extraClrScheme>`;
    })
    .join("");
  return `<a:extraClrSchemeLst>${items}</a:extraClrSchemeLst>`;
}

/** Parse a:extraClrSchemeLst. */
export function parseExtraColorSchemes(
  el: XmlElement | undefined,
  ctx: ReadContext,
): ExtraColorSchemeOptions[] | undefined {
  if (!el) return undefined;
  const result: ExtraColorSchemeOptions[] = [];
  for (const child of el.elements ?? []) {
    if (child.name !== "a:extraClrScheme") continue;
    const colorScheme = parseColorScheme(findChild(child, "a:clrScheme"), ctx);
    if (!colorScheme) continue;
    const colorMapping = parseColorMapping(findChild(child, "a:clrMap"));
    result.push(colorMapping ? { colorScheme, colorMapping } : { colorScheme });
  }
  return result.length > 0 ? result : undefined;
}
