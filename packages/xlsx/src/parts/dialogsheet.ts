/**
 * Dialogsheet types and descriptor for SpreadsheetML documents.
 *
 * A dialogsheet is a legacy Excel 5.0 dialog sheet (no cell data).
 *
 * Reference: OOXML transitional, sml.xsd, CT_Dialogsheet
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import type { UniversalMeasure } from "@office-open/core";
import { convertToInch } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attrs, attr, attrNum, escapeXml, findChild, stringifyElement } from "@office-open/xml";

// ── Types ──

export interface DialogsheetPageMargins {
  left?: number | UniversalMeasure;
  right?: number | UniversalMeasure;
  top?: number | UniversalMeasure;
  bottom?: number | UniversalMeasure;
  header?: number | UniversalMeasure;
  footer?: number | UniversalMeasure;
}

export interface DialogsheetPageSetup {
  /** Paper size (1=Letter, 9=A4, etc.) */
  paperSize?: number;
  /** Orientation ("default" | "portrait" | "landscape") */
  orientation?: string;
  /** Horizontal DPI */
  horizontalDpi?: number;
  /** Vertical DPI */
  verticalDpi?: number;
  /** Copies to print */
  copies?: number;
}

export interface DialogsheetProtectionOptions {
  /** Content is protected */
  content?: boolean;
  /** Objects are protected */
  objects?: boolean;
  /** Scenarios are protected */
  scenarios?: boolean;
}

export interface DialogsheetOptions {
  /** Sheet name */
  name?: string;
  /** Tab color (hex ARGB) */
  tabColor?: string;
  /** Published to a server (CT_SheetPr @published) */
  published?: boolean;
  /** VBA code name (CT_SheetPr @codeName) */
  codeName?: string;
  /** Page margins */
  pageMargins?: DialogsheetPageMargins;
  /** Page setup */
  pageSetup?: DialogsheetPageSetup;
  /** Sheet protection */
  sheetProtection?: DialogsheetProtectionOptions;
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
    if (opts.tabColor || opts.published || opts.codeName) {
      const prChildren: string[] = [];
      if (opts.tabColor) prChildren.push(`<tabColor${attrs({ rgb: opts.tabColor })}/>`);
      const prAttrs: string[] = [];
      if (opts.published) prAttrs.push(' published="1"');
      if (opts.codeName) prAttrs.push(` codeName="${escapeXml(opts.codeName)}"`);
      p.push(`<sheetPr${prAttrs.join("")}>${prChildren.join("")}</sheetPr>`);
    }

    // sheetProtection (optional)
    if (opts.sheetProtection) {
      const sp = opts.sheetProtection;
      const spAttrs: string[] = [];
      if (sp.content) spAttrs.push(` content="1"`);
      if (sp.objects) spAttrs.push(` objects="1"`);
      if (sp.scenarios) spAttrs.push(` scenarios="1"`);
      if (spAttrs.length > 0) {
        p.push(`<sheetProtection${spAttrs.join("")}/>`);
      }
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

    // pageSetup (optional)
    if (opts.pageSetup) {
      const ps = opts.pageSetup;
      p.push(
        `<pageSetup${attrs({
          paperSize: ps.paperSize,
          orientation: ps.orientation,
          horizontalDpi: ps.horizontalDpi,
          verticalDpi: ps.verticalDpi,
          copies: ps.copies,
        })}/>`,
      );
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
      if (parseOnOff(attr(sheetPrEl, "published"))) result.published = true;
      if (attr(sheetPrEl, "codeName")) result.codeName = attr(sheetPrEl, "codeName");
      const tcEl = findChild(sheetPrEl, "tabColor");
      if (tcEl && attr(tcEl, "rgb")) result.tabColor = attr(tcEl, "rgb");
    }

    // sheetProtection
    const spEl = findChild(el, "sheetProtection");
    if (spEl) {
      const sp: Partial<DialogsheetProtectionOptions> = {};
      if (parseOnOff(attr(spEl, "content"))) sp.content = true;
      if (parseOnOff(attr(spEl, "objects"))) sp.objects = true;
      if (parseOnOff(attr(spEl, "scenarios"))) sp.scenarios = true;
      result.sheetProtection = sp;
    }

    // pageMargins
    const pmEl = findChild(el, "pageMargins");
    if (pmEl) {
      const pm: Partial<DialogsheetPageMargins> = {};
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

    // pageSetup
    const psEl = findChild(el, "pageSetup");
    if (psEl) {
      const ps: Partial<DialogsheetPageSetup> = {};
      const pz = attrNum(psEl, "paperSize");
      if (pz !== undefined) ps.paperSize = pz;
      const orient = attr(psEl, "orientation");
      if (orient) ps.orientation = orient;
      const hdpi = attrNum(psEl, "horizontalDpi");
      if (hdpi !== undefined) ps.horizontalDpi = hdpi;
      const vdpi = attrNum(psEl, "verticalDpi");
      if (vdpi !== undefined) ps.verticalDpi = vdpi;
      const copies = attrNum(psEl, "copies");
      if (copies !== undefined) ps.copies = copies;
      result.pageSetup = ps;
    }

    // extLst — preserved verbatim
    const extEl = findChild(el, "extLst");
    if (extEl) result.extLst = stringifyElement(extEl);

    return result as DialogsheetOptions;
  },
};
