/**
 * Chartsheet types and descriptor for SpreadsheetML documents.
 *
 * A chartsheet is a worksheet that contains only a chart (no cells).
 *
 * Reference: OOXML transitional, sml.xsd, CT_Chartsheet
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import { derivePasswordHash } from "@office-open/core";
import type { PositiveUniversalMeasure } from "@office-open/core";
import { convertToInch } from "@office-open/core";
import type { ChartSpaceOptions } from "@office-open/core/chart";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import type { GraphicFrameLockingOptions } from "@office-open/core/drawing";
import { attrs, attr, attrMeasure, attrNum, escapeXml, findChild } from "@office-open/xml";
import { hashPassword } from "@util/index";

import { parseHeaderFooterEl } from "./worksheet/descriptor";
import { stringifyHeaderFooterXml } from "./worksheet/stringify";
import type { HeaderFooterOptions, PageMarginsOptions, PageOrientation } from "./worksheet/types";

// ── Types ──

/**
 * Chartsheet page setup (CT_CsPageSetup) — the chartsheet-specific variant of
 * pageSetup: paper size/orientation plus printer flags, no scale/fit fields.
 */
export interface ChartsheetPageSetup {
  /** Paper size (1=Letter, 9=A4, etc.) */
  paperSize?: number;
  /** Paper height (ST_PositiveUniversalMeasure) */
  paperHeight?: number | PositiveUniversalMeasure;
  /** Paper width (ST_PositiveUniversalMeasure) */
  paperWidth?: number | PositiveUniversalMeasure;
  /** First page number */
  firstPageNumber?: number;
  /** Orientation (ST_Orientation) */
  orientation?: PageOrientation;
  /** Use printer defaults (XSD default true — only false is emitted) */
  usePrinterDefaults?: boolean;
  /** Black and white printing */
  blackAndWhite?: boolean;
  /** Draft quality printing */
  draft?: boolean;
  /** Use firstPageNumber as the starting page number */
  useFirstPageNumber?: boolean;
  /** Horizontal DPI */
  horizontalDpi?: number;
  /** Vertical DPI */
  verticalDpi?: number;
  /** Copies to print */
  copies?: number;
}

export interface ChartsheetProtectionOptions {
  /**
   * Plain-text password — legacy Excel hash is computed automatically on
   * stringify. Authoring-only: not carried back by parse (see
   * SheetProtectionOptions); use the algorithmName quadruplet for round-trip.
   */
  password?: string;
  /** Modern encryption: algorithm name (e.g. "SHA-512") */
  algorithmName?: string;
  /** Modern encryption: base64-encoded hash value */
  hashValue?: string;
  /** Modern encryption: base64-encoded salt value */
  saltValue?: string;
  /** Modern encryption: spin count for hash iteration */
  spinCount?: number;
  /** Content is protected */
  content?: boolean;
  /** Objects are protected */
  objects?: boolean;
}

export interface ChartsheetOptions {
  /** Sheet name */
  name?: string;
  /** Tab color (hex ARGB, e.g. "FF4472C4") */
  tabColor?: string;
  /** Page margins */
  pageMargins?: PageMarginsOptions;
  /** Page setup */
  pageSetup?: ChartsheetPageSetup;
  /** Header/footer */
  headerFooter?: HeaderFooterOptions;
  /** Sheet protection */
  sheetProtection?: ChartsheetProtectionOptions;
  /** Published to server (CT_ChartsheetPr `@published`, XSD default true — only false is emitted) */
  published?: boolean;
  /** VBA code name (CT_ChartsheetPr `@codeName`) */
  codeName?: string;
  /** Zoom to fit (CT_ChartsheetView `@zoomToFit`) */
  zoomToFit?: boolean;
  /** Chart definition — the shared chart-space model, same shape as a worksheet chart. */
  chart?: ChartSpaceOptions;
  /** Macro reference on the chart's graphicFrame (CT_GraphicFrame/@macro). */
  macro?: string;
  /** Frame locks on the chart's graphicFrame (a:graphicFrameLocks). */
  frameLocks?: GraphicFrameLockingOptions;
}

// ── Descriptor Types ──

export interface ChartsheetDescriptorOptions extends ChartsheetOptions {
  /** Relationship ID for the drawing (set by compiler) */
  drawingRId: string;
}

// ── Descriptor ──

export const chartsheetDesc: CustomDescriptor<ChartsheetDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const p: string[] = [
      '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    ];

    // sheetPr (optional)
    if (opts.tabColor || opts.published !== undefined || opts.codeName) {
      const prAttrs: string[] = [];
      if (opts.tabColor) prAttrs.push(`<tabColor${attrs({ rgb: opts.tabColor })}/>`);
      const spAttrs: string[] = [];
      // XSD default true — emit only the explicit-false form (0).
      if (opts.published === false) spAttrs.push(' published="0"');
      if (opts.codeName) spAttrs.push(` codeName="${escapeXml(opts.codeName)}"`);
      p.push(`<sheetPr${spAttrs.join("")}>${prAttrs.join("")}</sheetPr>`);
    }

    // sheetViews (required)
    const svAttrs: string[] = ['workbookViewId="0"'];
    if (opts.zoomToFit) svAttrs.push('zoomToFit="1"');
    p.push(`<sheetViews><sheetView ${svAttrs.join(" ")}/></sheetViews>`);

    // sheetProtection (optional) — CT_ChartsheetProtection
    if (opts.sheetProtection) {
      const sp = opts.sheetProtection;
      const spAttrs: Record<string, string | number | boolean | undefined> = {};
      if (sp.password) spAttrs.password = hashPassword(sp.password);
      let derived: ReturnType<typeof derivePasswordHash> | undefined;
      if (sp.password !== undefined && sp.hashValue === undefined) {
        derived = derivePasswordHash(sp.password);
      }
      spAttrs.algorithmName = sp.algorithmName ?? derived?.algorithmName;
      spAttrs.hashValue = sp.hashValue ?? derived?.hashValue;
      spAttrs.saltValue = sp.saltValue ?? derived?.saltValue;
      if (sp.spinCount !== undefined) spAttrs.spinCount = sp.spinCount;
      else if (derived) spAttrs.spinCount = derived.spinCount;
      if (sp.content) spAttrs.content = 1;
      if (sp.objects) spAttrs.objects = 1;
      if (Object.keys(spAttrs).length > 0) {
        p.push(`<sheetProtection${attrs(spAttrs)}/>`);
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

    // pageSetup (optional) — CT_CsPageSetup
    if (opts.pageSetup) {
      const ps = opts.pageSetup;
      const psAttrs: Record<string, string | number | boolean | undefined> = {};
      if (ps.paperSize !== undefined) psAttrs.paperSize = ps.paperSize;
      // ST_PositiveUniversalMeasure requires a unit suffix; a bare number means mm.
      if (ps.paperHeight !== undefined)
        psAttrs.paperHeight =
          typeof ps.paperHeight === "number" ? `${ps.paperHeight}mm` : ps.paperHeight;
      if (ps.paperWidth !== undefined)
        psAttrs.paperWidth =
          typeof ps.paperWidth === "number" ? `${ps.paperWidth}mm` : ps.paperWidth;
      if (ps.firstPageNumber !== undefined) psAttrs.firstPageNumber = ps.firstPageNumber;
      if (ps.orientation && ps.orientation !== "default") psAttrs.orientation = ps.orientation;
      // XSD default true — emit only the explicit-false form (0).
      if (ps.usePrinterDefaults === false) psAttrs.usePrinterDefaults = 0;
      if (ps.blackAndWhite) psAttrs.blackAndWhite = 1;
      if (ps.draft) psAttrs.draft = 1;
      if (ps.useFirstPageNumber) psAttrs.useFirstPageNumber = 1;
      if (ps.horizontalDpi !== undefined) psAttrs.horizontalDpi = ps.horizontalDpi;
      if (ps.verticalDpi !== undefined) psAttrs.verticalDpi = ps.verticalDpi;
      if (ps.copies !== undefined) psAttrs.copies = ps.copies;
      p.push(`<pageSetup${attrs(psAttrs)}/>`);
    }

    // headerFooter (optional) — CT_HeaderFooter, shared with worksheet
    if (opts.headerFooter) {
      const hfXml = stringifyHeaderFooterXml(opts.headerFooter);
      if (hfXml) p.push(hfXml);
    }

    // drawing (required)
    p.push(`<drawing r:id="${escapeXml(opts.drawingRId)}"/>`);

    p.push("</chartsheet>");
    return p.join("");
  },

  parse(el, _ctx) {
    const result: Partial<ChartsheetDescriptorOptions> = {};

    // nativeTypeAttributes (xlsx parse path) coerces "1"/"0" to numbers, so
    // boolean attribute checks use String() coercion.
    // sheetPr
    const sheetPr = findChild(el, "sheetPr");
    if (sheetPr) {
      // XSD default true — only the explicit "0" carries information back.
      if (String(attr(sheetPr, "published")) === "0") result.published = false;
      if (attr(sheetPr, "codeName")) result.codeName = attr(sheetPr, "codeName");
      const tabColor = findChild(sheetPr, "tabColor");
      if (tabColor) {
        const rgb = attr(tabColor, "rgb");
        if (rgb) result.tabColor = String(rgb);
      }
    }

    // sheetViews
    const sheetViews = findChild(el, "sheetViews");
    if (sheetViews) {
      const sv = findChild(sheetViews, "sheetView");
      if (sv && parseOnOff(attr(sv, "zoomToFit"))) result.zoomToFit = true;
    }

    // sheetProtection — CT_ChartsheetProtection
    const sheetProtectionEl = findChild(el, "sheetProtection");
    if (sheetProtectionEl) {
      const sp: Partial<ChartsheetProtectionOptions> = {};
      // @password (legacy hash) not read back — see parseSheetProtectionEl note.
      const an = attr(sheetProtectionEl, "algorithmName");
      if (an) sp.algorithmName = an;
      const hv = attr(sheetProtectionEl, "hashValue");
      if (hv) sp.hashValue = hv;
      const sv = attr(sheetProtectionEl, "saltValue");
      if (sv) sp.saltValue = sv;
      const sc = attrNum(sheetProtectionEl, "spinCount");
      if (sc !== undefined) sp.spinCount = sc;
      if (parseOnOff(attr(sheetProtectionEl, "content"))) sp.content = true;
      if (parseOnOff(attr(sheetProtectionEl, "objects"))) sp.objects = true;
      result.sheetProtection = sp;
    }

    // pageMargins
    const pageMarginsEl = findChild(el, "pageMargins");
    if (pageMarginsEl) {
      const pm: Partial<PageMarginsOptions> = {};
      const ml = attrNum(pageMarginsEl, "left");
      if (ml !== undefined) pm.left = ml;
      const mr = attrNum(pageMarginsEl, "right");
      if (mr !== undefined) pm.right = mr;
      const mt = attrNum(pageMarginsEl, "top");
      if (mt !== undefined) pm.top = mt;
      const mb = attrNum(pageMarginsEl, "bottom");
      if (mb !== undefined) pm.bottom = mb;
      const mh = attrNum(pageMarginsEl, "header");
      if (mh !== undefined) pm.header = mh;
      const mf = attrNum(pageMarginsEl, "footer");
      if (mf !== undefined) pm.footer = mf;
      result.pageMargins = pm;
    }

    // pageSetup — CT_CsPageSetup
    const pageSetupEl = findChild(el, "pageSetup");
    if (pageSetupEl) {
      const ps: Partial<ChartsheetPageSetup> = {};
      const pz = attrNum(pageSetupEl, "paperSize");
      if (pz !== undefined) ps.paperSize = pz;
      const ph = attrMeasure(pageSetupEl, "paperHeight");
      if (ph !== undefined) ps.paperHeight = ph as number | PositiveUniversalMeasure;
      const pw = attrMeasure(pageSetupEl, "paperWidth");
      if (pw !== undefined) ps.paperWidth = pw as number | PositiveUniversalMeasure;
      const fpn = attrNum(pageSetupEl, "firstPageNumber");
      if (fpn !== undefined) ps.firstPageNumber = fpn;
      const orient = attr(pageSetupEl, "orientation");
      if (orient) ps.orientation = orient as PageOrientation;
      // XSD default true — only the explicit "0" carries information back.
      if (String(attr(pageSetupEl, "usePrinterDefaults")) === "0") ps.usePrinterDefaults = false;
      if (parseOnOff(attr(pageSetupEl, "blackAndWhite"))) ps.blackAndWhite = true;
      if (parseOnOff(attr(pageSetupEl, "draft"))) ps.draft = true;
      if (parseOnOff(attr(pageSetupEl, "useFirstPageNumber"))) ps.useFirstPageNumber = true;
      const hdpi = attrNum(pageSetupEl, "horizontalDpi");
      if (hdpi !== undefined) ps.horizontalDpi = hdpi;
      const vdpi = attrNum(pageSetupEl, "verticalDpi");
      if (vdpi !== undefined) ps.verticalDpi = vdpi;
      const copies = attrNum(pageSetupEl, "copies");
      if (copies !== undefined) ps.copies = copies;
      result.pageSetup = ps;
    }

    // headerFooter — CT_HeaderFooter, shared with worksheet
    const headerFooterEl = findChild(el, "headerFooter");
    if (headerFooterEl) {
      result.headerFooter = parseHeaderFooterEl(headerFooterEl);
    }

    return result as ChartsheetDescriptorOptions;
  },
};
