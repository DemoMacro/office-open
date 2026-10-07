/**
 * Worksheet — parse helpers shared by the descriptor.
 *
 * @module
 */
import { parseOnOff } from "@office-open/core";
import { attr, attrNum } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import { parseColorOptions, type ColorOptions } from "../../shared/color";
import type { CfvoOptions, CfvoType, PageBreakOptions } from "./types";

/**
 * Parse a conditional-formatting `<color>` (CT_Color). Every channel is an
 * alternative — rgb hex, a theme slot (with optional tint), or a legacy
 * palette index. Returns undefined when the element carries none of them.
 */
export function parseCfColor(el: XmlElement): ColorOptions | undefined {
  const result = parseColorOptions(el);
  return Object.keys(result).length > 0 ? result : undefined;
}

/** Parse a CT_PageBreak (w:rowBreaks / w:colBreaks) into break entries. */
export function parsePageBreaks(el: XmlElement): PageBreakOptions[] {
  const breaks: PageBreakOptions[] = [];
  for (const brkEl of el.elements ?? []) {
    if (brkEl.name !== "brk") continue;
    const id = attrNum(brkEl, "id");
    if (id === undefined) continue;
    const b: PageBreakOptions = { id };
    const min = attrNum(brkEl, "min");
    if (min !== undefined) b.min = min;
    const max = attrNum(brkEl, "max");
    if (max !== undefined) b.max = max;
    if (parseOnOff(attr(brkEl, "man"))) b.manual = true;
    if (parseOnOff(attr(brkEl, "pt"))) b.pivot = true;
    breaks.push(b);
  }
  return breaks;
}

export function parseCfvo(el: XmlElement): CfvoOptions {
  const result: CfvoOptions = { type: (attr(el, "type") ?? "num") as CfvoType };
  const val = attr(el, "val");
  if (val !== undefined) result.val = isNaN(Number(val)) ? val : Number(val);
  const gte = attr(el, "gte");
  if (gte !== undefined) result.gteRaw = gte;
  if (String(attr(el, "gte")) === "0") result.gte = false;
  return result;
}
