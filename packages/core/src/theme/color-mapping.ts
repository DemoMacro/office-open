/**
 * Color mapping (a:clrMap / CT_ColorMapping) stringify + parse.
 *
 * Remaps the 12 scheme slots (bg1/tx1/bg2/tx2/accent1-6/hlink/folHlink) to
 * ST_ColorSchemeIndex tokens (dk1/lt1/dk2/lt2/accent1-6/hlink/folHlink).
 *
 * @module
 */
import type { Element as XmlElement } from "@office-open/xml";

import type { ColorMappingOptions } from "./theme-options";

/** clrMap attribute token → ColorMappingOptions key. */
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

/** Serialize a:clrMap. */
export function stringifyColorMapping(opts: ColorMappingOptions): string {
  const attrs = MAPPING_ATTRS.map(({ attr, key }) => `${attr}="${opts[key]}"`).join(" ");
  return `<a:clrMap ${attrs}/>`;
}

/** Parse a:clrMap. */
export function parseColorMapping(el: XmlElement | undefined): ColorMappingOptions | undefined {
  if (!el?.attributes) return undefined;
  const result = {} as ColorMappingOptions;
  let hasAny = false;
  for (const { attr, key } of MAPPING_ATTRS) {
    const value = el.attributes[attr];
    if (value !== undefined) {
      result[key] = String(value);
      hasAny = true;
    }
  }
  return hasAny ? result : undefined;
}
