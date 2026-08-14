/**
 * Color mapping (a:clrMap / CT_ColorMapping) stringify + parse.
 *
 * Remaps the 12 semantic color slots to theme color slots. Public options use
 * full words; this module owns the abbreviated OOXML tokens.
 *
 * @module
 */
import { escapeXml } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type { ColorMappingOptions, ColorSchemeIndex } from "./theme-options";

const MAPPING_ATTRS: ReadonlyArray<{ attr: string; key: keyof ColorMappingOptions }> = [
  { attr: "bg1", key: "background1" },
  { attr: "tx1", key: "text1" },
  { attr: "bg2", key: "background2" },
  { attr: "tx2", key: "text2" },
  { attr: "accent1", key: "accent1" },
  { attr: "accent2", key: "accent2" },
  { attr: "accent3", key: "accent3" },
  { attr: "accent4", key: "accent4" },
  { attr: "accent5", key: "accent5" },
  { attr: "accent6", key: "accent6" },
  { attr: "hlink", key: "hyperlink" },
  { attr: "folHlink", key: "followedHyperlink" },
];

const COLOR_SCHEME_INDEX_TO_XML: Record<ColorSchemeIndex, string> = {
  dark1: "dk1",
  light1: "lt1",
  dark2: "dk2",
  light2: "lt2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hyperlink: "hlink",
  followedHyperlink: "folHlink",
};

const XML_TO_COLOR_SCHEME_INDEX = Object.fromEntries(
  Object.entries(COLOR_SCHEME_INDEX_TO_XML).map(([key, value]) => [value, key]),
) as Record<string, ColorSchemeIndex>;

/** Standard Office mapping from semantic slots to theme color slots. */
export const DEFAULT_COLOR_MAPPING: ColorMappingOptions = {
  background1: "light1",
  text1: "dark1",
  background2: "light2",
  text2: "dark2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hyperlink: "hyperlink",
  followedHyperlink: "followedHyperlink",
};

/** Serialize a complete, XSD-valid CT_ColorMapping element. */
export function stringifyColorMapping(
  opts: Partial<ColorMappingOptions> | undefined,
  elementName = "a:clrMap",
): string {
  const merged = { ...DEFAULT_COLOR_MAPPING, ...opts };
  const attrs = MAPPING_ATTRS.map(({ attr, key }) => {
    const value = COLOR_SCHEME_INDEX_TO_XML[merged[key]];
    return `${attr}="${escapeXml(value)}"`;
  }).join(" ");
  return `<${elementName} ${attrs}/>`;
}

/** Parse CT_ColorMapping attributes into full-word public keys and values. */
export function parseColorMapping(
  el: XmlElement | undefined,
): Partial<ColorMappingOptions> | undefined {
  if (!el?.attributes) return undefined;
  const result: Partial<ColorMappingOptions> = {};
  for (const { attr, key } of MAPPING_ATTRS) {
    const value = el.attributes[attr];
    if (value === undefined) continue;
    const parsed = XML_TO_COLOR_SCHEME_INDEX[String(value)];
    if (parsed !== undefined) result[key] = parsed;
  }
  return Object.keys(result).length > 0 ? result : undefined;
}
