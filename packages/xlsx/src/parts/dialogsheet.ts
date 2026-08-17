/**
 * Dialogsheet types and descriptor for SpreadsheetML documents.
 *
 * A dialogsheet is a legacy Excel 5.0 dialog sheet (no cell data).
 *
 * Reference: OOXML transitional, sml.xsd, CT_Dialogsheet
 *
 * @module
 */

import { convertToInch } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attrs, attr, attrNum, escapeXml, findChild, stringifyElement } from "@office-open/xml";

import { parsePageSetupEl, parseSheetProtectionEl } from "./worksheet/descriptor";
import { stringifyPageSetupXml, stringifySheetProtectionXml } from "./worksheet/stringify";
import type {
  PageMarginsOptions,
  PageSetupOptions,
  SheetProtectionOptions,
} from "./worksheet/types";

// ── Types ──

export interface DialogsheetOptions {
  /** Sheet name */
  name?: string;
  /** Tab color (hex ARGB) */
  tabColor?: string;
  /** Published to a server (CT_SheetPr `@published`, XSD default true — only false is emitted) */
  published?: boolean;
  /** VBA code name (CT_SheetPr `@codeName`) */
  codeName?: string;
  /** Page margins */
  pageMargins?: PageMarginsOptions;
  /** Page setup */
  pageSetup?: PageSetupOptions;
  /** Sheet protection */
  sheetProtection?: SheetProtectionOptions;
  /** Raw extension list preserved verbatim (extLst) */
  extLst?: string;
}

// ── Descriptor ──

export const dialogsheetDesc: CustomDescriptor<DialogsheetOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const p: string[] = [
      '<dialogsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    ];

    // sheetPr (optional)
    if (opts.tabColor || opts.published !== undefined || opts.codeName) {
      const prChildren: string[] = [];
      if (opts.tabColor) prChildren.push(`<tabColor${attrs({ rgb: opts.tabColor })}/>`);
      const prAttrs: string[] = [];
      // XSD default true — emit only the explicit-false form (0).
      if (opts.published === false) prAttrs.push(' published="0"');
      if (opts.codeName) prAttrs.push(` codeName="${escapeXml(opts.codeName)}"`);
      p.push(`<sheetPr${prAttrs.join("")}>${prChildren.join("")}</sheetPr>`);
    }

    // sheetProtection (optional) — CT_SheetProtection, shared with worksheet
    if (opts.sheetProtection) {
      p.push(stringifySheetProtectionXml(opts.sheetProtection));
    }

    // pageMargins (optional)
    if (opts.pageMargins) {
      const pm = opts.pageMargins;
      p.push(
        `<pageMargins${attrs({
          left: convertToInch(pm.left ?? 0.7),
          right: convertToInch(pm.right ?? 0.7),
          top: convertToInch(pm.top ?? 0.75),
          bottom: convertToInch(pm.bottom ?? 0.75),
          header: convertToInch(pm.header ?? 0.3),
          footer: convertToInch(pm.footer ?? 0.3),
        })}/>`,
      );
    }

    // pageSetup (optional) — CT_PageSetup, shared with worksheet
    if (opts.pageSetup) {
      p.push(stringifyPageSetupXml(opts.pageSetup));
    }

    if (opts.extLst) p.push(opts.extLst);

    p.push("</dialogsheet>");
    return p.join("");
  },

  parse(el, _ctx) {
    const result: Partial<DialogsheetOptions> = {};

    // sheetPr
    const sheetPrEl = findChild(el, "sheetPr");
    if (sheetPrEl) {
      // XSD default true — only the explicit "0" carries information back.
      if (String(attr(sheetPrEl, "published")) === "0") result.published = false;
      if (attr(sheetPrEl, "codeName")) result.codeName = attr(sheetPrEl, "codeName");
      const tcEl = findChild(sheetPrEl, "tabColor");
      if (tcEl && attr(tcEl, "rgb")) result.tabColor = attr(tcEl, "rgb");
    }

    // sheetProtection — CT_SheetProtection, shared with worksheet
    const spEl = findChild(el, "sheetProtection");
    if (spEl?.attributes) {
      result.sheetProtection = parseSheetProtectionEl(spEl);
    }

    // pageMargins
    const pmEl = findChild(el, "pageMargins");
    if (pmEl) {
      const pm: Partial<PageMarginsOptions> = {};
      const ml = attrNum(pmEl, "left");
      if (ml !== undefined) pm.left = ml;
      const mr = attrNum(pmEl, "right");
      if (mr !== undefined) pm.right = mr;
      const mt = attrNum(pmEl, "top");
      if (mt !== undefined) pm.top = mt;
      const mb = attrNum(pmEl, "bottom");
      if (mb !== undefined) pm.bottom = mb;
      const mh = attrNum(pmEl, "header");
      if (mh !== undefined) pm.header = mh;
      const mf = attrNum(pmEl, "footer");
      if (mf !== undefined) pm.footer = mf;
      result.pageMargins = pm;
    }

    // pageSetup — CT_PageSetup, shared with worksheet
    const psEl = findChild(el, "pageSetup");
    if (psEl) {
      result.pageSetup = parsePageSetupEl(psEl);
    }

    // extLst — preserved verbatim
    const extEl = findChild(el, "extLst");
    if (extEl) result.extLst = stringifyElement(extEl);

    return result as DialogsheetOptions;
  },
};
