/**
 * Calculation Chain types and descriptor.
 *
 * Reference: OOXML transitional, sml.xsd, CT_CalcChain / CT_CalcCell
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attrs } from "@office-open/xml";

// ── Types ──

/** One calc-chain entry (xl/calcChain.xml c:r — a formula cell in calc order). */
export interface CalcCell {
  /** Cell reference, e.g. "A1" */
  reference: string;
  /** Sheet index (1-based); omitted cells inherit the previous entry's index */
  sheetIndex?: number;
  /** Array formula */
  array?: boolean;
  /** Child chain — calculations farmed out to another thread (CT_CalcCell `@s`) */
  childChain?: boolean;
  /** Calculation level increased relative to the previous entry (CT_CalcCell `@l`) */
  newLevel?: boolean;
  /** Calculation moved to a new dependency thread (CT_CalcCell `@t`) */
  newThread?: boolean;
}

export interface CalcChainOptions {
  cells: CalcCell[];
}

// ── Descriptor ──

export const calcChainDesc: CustomDescriptor<CalcChainOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const parts: string[] = [
      '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    ];
    for (const cell of opts.cells) {
      const cellAttrs: Record<string, string | number | boolean> = {
        r: cell.reference,
      };
      if (cell.sheetIndex !== undefined) cellAttrs.i = cell.sheetIndex;
      if (cell.array !== undefined) cellAttrs.a = cell.array;
      if (cell.childChain !== undefined) cellAttrs.s = cell.childChain;
      if (cell.newLevel !== undefined) cellAttrs.l = cell.newLevel;
      if (cell.newThread !== undefined) cellAttrs.t = cell.newThread;
      parts.push(`<c${attrs(cellAttrs)}/>`);
    }
    parts.push("</calcChain>");
    return parts.join("");
  },

  parse(el, _ctx) {
    const result: Partial<CalcChainOptions> = {};
    const cells: CalcCell[] = [];
    for (const child of el.elements ?? []) {
      if (child.name !== "c") continue;
      const r = child.attributes?.["r"];
      if (r) {
        const cell: CalcCell = {
          reference: String(r),
        };
        const i = child.attributes?.["i"];
        if (i !== undefined) cell.sheetIndex = Number(i);
        if (child.attributes?.["a"] !== undefined)
          cell.array = parseOnOff(child.attributes["a"]) ?? true;
        if (child.attributes?.["s"] !== undefined)
          cell.childChain = parseOnOff(child.attributes["s"]) ?? true;
        if (child.attributes?.["l"] !== undefined)
          cell.newLevel = parseOnOff(child.attributes["l"]) ?? true;
        if (child.attributes?.["t"] !== undefined)
          cell.newThread = parseOnOff(child.attributes["t"]) ?? true;
        cells.push(cell);
      }
    }
    result.cells = cells;
    return result as CalcChainOptions;
  },
};
