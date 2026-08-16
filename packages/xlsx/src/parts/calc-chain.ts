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

export interface CalcCell {
  /** Cell reference, e.g. "A1" */
  reference: string;
  /** Sheet index (1-based) */
  sheetIndex: number;
  /** Array formula */
  array?: boolean;
  /** Child chain — calculations farmed out to another thread (CT_CalcCell `@l`) */
  childChain?: boolean;
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
        i: cell.sheetIndex,
      };
      if (cell.array) cellAttrs.a = true;
      if (cell.childChain) cellAttrs.l = true;
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
      const i = child.attributes?.["i"];
      if (r && i) {
        const cell: CalcCell = {
          reference: String(r),
          sheetIndex: Number(i),
        };
        if (child.attributes?.["a"]) cell.array = true;
        if (parseOnOff(child.attributes?.["l"])) cell.childChain = true;
        cells.push(cell);
      }
    }
    result.cells = cells;
    return result as CalcChainOptions;
  },
};
