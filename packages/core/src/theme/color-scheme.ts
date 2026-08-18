/**
 * Color scheme (a:clrScheme / CT_ColorScheme) stringify + parse.
 *
 * Every slot accepts a hex string (emits a:srgbClr) or a structured
 * SystemColorOptions (emits a:sysClr verbatim). dk1/lt1 default to the
 * conventional windowText/window sysClr form; parse reads both forms back
 * losslessly, so a source's sysClr spelling survives round-trip.
 *
 * @module
 */
import { findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type { ReadContext } from "../descriptor";
import { systemColorDesc } from "../drawing";
import type { SystemColorOptions } from "../drawing";
import { DEFAULT_COLORS } from "./default-colors";
import type { ColorSchemeOptions } from "./theme-options";

export type { SchemeColorValue } from "./theme-options";

/** Color keys excluding the clrScheme/`@name` attribute. */
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

/** The conventional dk1/lt1 sysClr forms emitted when a slot is unset. */
const DEFAULT_SYS_CLR: Partial<Record<ColorKey, string>> = {
  dark1: "windowText",
  light1: "window",
};

function stringifyColorTag(
  tag: string,
  key: ColorKey,
  opts: ColorSchemeOptions | undefined,
): string {
  const userValue = opts?.[key];
  if (userValue !== undefined) {
    const colorXml =
      typeof userValue === "string"
        ? `<a:srgbClr val="${userValue}"/>`
        : (systemColorDesc.stringify(userValue, undefined as never) ?? "");
    return `<${tag}>${colorXml}</${tag}>`;
  }
  const defaultValue = DEFAULT_COLORS[key];
  const sysClr = DEFAULT_SYS_CLR[key];
  if (sysClr) {
    return `<${tag}><a:sysClr val="${sysClr}" lastClr="${defaultValue}"/></${tag}>`;
  }
  return `<${tag}><a:srgbClr val="${defaultValue}"/></${tag}>`;
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
export function parseColorScheme(
  el: XmlElement | undefined,
  ctx: ReadContext,
): ColorSchemeOptions | undefined {
  if (!el) return undefined;
  const result: Partial<ColorSchemeOptions> = {};
  const name = el.attributes?.["name"];
  if (name) result.name = String(name);
  for (const { tag, key } of COLOR_TAGS) {
    const colorEl = findChild(el, tag);
    if (!colorEl) continue;
    const sysClr = findChild(colorEl, "a:sysClr");
    if (sysClr) {
      result[key] = systemColorDesc.parse(sysClr, ctx) as SystemColorOptions;
      continue;
    }
    const srgb = findChild(colorEl, "a:srgbClr");
    const val = srgb?.attributes?.["val"];
    if (val !== undefined) result[key] = String(val);
  }
  return result;
}
