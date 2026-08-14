/**
 * Worksheet — parse helpers shared by the descriptor.
 *
 * @module
 */
import { parseOnOff } from "@office-open/core";
import { attr, attrNum } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type { CfvoOptions, CfvoType, PageBreakOptions } from "./types";

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
  if (String(attr(el, "gte")) === "0") result.gte = false;
  return result;
}

export function parseCellRef(ref: string): { row: number; col: number } | undefined {
  const match = ref.match(/^([A-Z]+)(\d+)$/);
  if (!match) return undefined;
  const colStr = match[1] ?? "";
  const row = parseInt(match[2] ?? "0", 10);
  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }
  return { row, col };
}
