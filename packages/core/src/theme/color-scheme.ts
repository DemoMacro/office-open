/**
 * Color scheme (a:clrScheme / CT_ColorScheme) stringify + parse.
 *
 * dk1/lt1 use a:sysClr (windowText/window) by Office convention; only when the
 * caller explicitly supplies a value do they switch to a:srgbClr. parse reads
 * back the effective color value (srgbClr val or sysClr lastClr).
 *
 * @module
 */
import { findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import { DEFAULT_COLORS } from "./default-colors";
import type { ColorSchemeOptions } from "./theme-options";

/** Color keys excluding the clrScheme/@name attribute. */
type ColorKey = Exclude<keyof ColorSchemeOptions, "name">;

/** clrScheme child tag → ColorSchemeOptions key, in schema order. */
const COLOR_TAGS: ReadonlyArray<{ tag: string; key: ColorKey }> = [
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

/** Read the effective hex color from a color element (srgbClr val or sysClr lastClr). */
function readColorValue(el: XmlElement | undefined): string | undefined {
  if (!el) return undefined;
  const srgb = findChild(el, "a:srgbClr");
  if (srgb) return String(srgb.attributes?.["val"] ?? "");
  const sysClr = findChild(el, "a:sysClr");
  if (sysClr) return String(sysClr.attributes?.["lastClr"] ?? "");
  return undefined;
}

/** dk1/lt1 emit sysClr by default, srgbClr when explicitly supplied. */
function stringifyColorTag(
  tag: string,
  key: ColorKey,
  opts: ColorSchemeOptions | undefined,
): string {
  const userValue = opts?.[key];
  const defaultValue = DEFAULT_COLORS[key];
  if (key === "dark1") {
    return userValue !== undefined
      ? `<${tag}><a:srgbClr val="${userValue}"/></${tag}>`
      : `<${tag}><a:sysClr val="windowText" lastClr="${defaultValue}"/></${tag}>`;
  }
  if (key === "light1") {
    return userValue !== undefined
      ? `<${tag}><a:srgbClr val="${userValue}"/></${tag}>`
      : `<${tag}><a:sysClr val="window" lastClr="${defaultValue}"/></${tag}>`;
  }
  const value = userValue ?? defaultValue;
  return `<${tag}><a:srgbClr val="${value}"/></${tag}>`;
}

/** Serialize a:clrScheme. Undefined options emit the full default scheme. */
export function stringifyColorScheme(
  opts: ColorSchemeOptions | undefined,
  fallbackName: string,
): string {
  const name = opts?.name ?? fallbackName;
  const tags = COLOR_TAGS.map(({ tag, key }) => stringifyColorTag(tag, key, opts)).join("");
  return `<a:clrScheme name="${name}">${tags}</a:clrScheme>`;
}

/** Parse a:clrScheme into the subset of colors explicitly present. */
export function parseColorScheme(el: XmlElement | undefined): ColorSchemeOptions | undefined {
  if (!el) return undefined;
  const result: Partial<ColorSchemeOptions> = {};
  const name = el.attributes?.["name"];
  if (name) result.name = String(name);
  for (const { tag, key } of COLOR_TAGS) {
    const value = readColorValue(findChild(el, tag));
    if (value) result[key] = value;
  }
  return result;
}
