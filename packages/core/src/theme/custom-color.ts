/**
 * Custom colors (a:custClr / CT_CustomColor) stringify + parse.
 *
 * Each entry is a named color choice; the list wrapper (a:custClrLst /
 * CT_CustomColorList) sits after extraClrSchemeLst in a theme.
 *
 * @module
 */
import type { Element as XmlElement } from "@office-open/xml";

import type { ReadContext, WriteContext } from "../descriptor";
import { parseColorChoice, stringifyColorChoice } from "../drawing";
import type { SolidFillOptions } from "../drawing";
import type { CustomColorOptions } from "./theme-options";

/** Serialize a:custClrLst. Returns "" when no custom colors set (optional element). */
export function stringifyCustomColors(
  opts: CustomColorOptions[] | undefined,
  ctx: WriteContext,
): string {
  if (!opts || opts.length === 0) return "";
  const items = opts
    .map((c) => {
      const name = c.name ? ` name="${c.name}"` : "";
      return `<a:custClr${name}>${stringifyColorChoice(c.color, ctx)}</a:custClr>`;
    })
    .join("");
  return `<a:custClrLst>${items}</a:custClrLst>`;
}

/** Parse a:custClrLst. */
export function parseCustomColors(
  el: XmlElement | undefined,
  ctx: ReadContext,
): CustomColorOptions[] | undefined {
  if (!el) return undefined;
  const result: CustomColorOptions[] = [];
  for (const child of el.elements ?? []) {
    if (child.name !== "a:custClr") continue;
    const color: SolidFillOptions | undefined = parseColorChoice(child, ctx);
    if (!color) continue;
    const name = child.attributes?.["name"];
    result.push(name !== undefined ? { name: String(name), color } : { color });
  }
  return result.length > 0 ? result : undefined;
}
