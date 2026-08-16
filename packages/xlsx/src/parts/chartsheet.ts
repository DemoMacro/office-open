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
import type { UniversalMeasure } from "@office-open/core";
import { convertToInch } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attrs, attr, attrNum, escapeXml, findChild, textOf } from "@office-open/xml";

// ── Types ──

export interface ChartsheetPageMargins {
  left?: number | UniversalMeasure;
  right?: number | UniversalMeasure;
  top?: number | UniversalMeasure;
  bottom?: number | UniversalMeasure;
  header?: number | UniversalMeasure;
  footer?: number | UniversalMeasure;
}

export interface ChartsheetPageSetup {
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

export interface ChartsheetProtectionOptions {
  /** Content is protected */
  content?: boolean;
  /** Objects are protected */
  objects?: boolean;
}

export interface ChartsheetHeaderFooterOptions {
  /** Different first page header/footer */
  differentFirst?: boolean;
  /** Different odd/even page headers/footers */
  differentOddEven?: boolean;
  /** Odd page header */
  oddHeader?: string;
  /** Odd page footer */
  oddFooter?: string;
}

export interface ChartsheetOptions {
  /** Sheet name */
  name?: string;
  /** Tab color (hex ARGB, e.g. "FF4472C4") */
  tabColor?: string;
  /** Page margins */
  pageMargins?: ChartsheetPageMargins;
  /** Page setup */
  pageSetup?: ChartsheetPageSetup;
  /** Header/footer */
  headerFooter?: ChartsheetHeaderFooterOptions;
  /** Sheet protection */
  sheetProtection?: ChartsheetProtectionOptions;
  /** Published to server (CT_ChartsheetPr `@published`) */
  published?: boolean;
  /** VBA code name (CT_ChartsheetPr `@codeName`) */
  codeName?: string;
  /** Zoom to fit (CT_ChartsheetView `@zoomToFit`) */
  zoomToFit?: boolean;
  /** Chart definition (type, title, series, etc.) */
  chart: {
    type: string;
    title?: string;
    categories?: string[];
    series: {
      name: string;
      values: number[];
    }[];
  };
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
    if (opts.tabColor || opts.published || opts.codeName) {
      const prAttrs: string[] = [];
      if (opts.tabColor) prAttrs.push(`<tabColor${attrs({ rgb: opts.tabColor })}/>`);
      const spAttrs: string[] = [];
      if (opts.published) spAttrs.push(' published="1"');
      if (opts.codeName) spAttrs.push(` codeName="${escapeXml(opts.codeName)}"`);
      p.push(`<sheetPr${spAttrs.join("")}>${prAttrs.join("")}</sheetPr>`);
    }

    // sheetViews (required)
    const svAttrs: string[] = ['workbookViewId="0"'];
    if (opts.zoomToFit) svAttrs.push('zoomToFit="1"');
    p.push(`<sheetViews><sheetView ${svAttrs.join(" ")}/></sheetViews>`);

    // sheetProtection (optional)
    if (opts.sheetProtection) {
      const sp = opts.sheetProtection;
      const spAttrs: string[] = [];
      if (sp.content) spAttrs.push(` content="1"`);
      if (sp.objects) spAttrs.push(` objects="1"`);
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

    // headerFooter (optional)
    if (opts.headerFooter) {
      const hf = opts.headerFooter;
      const hfParts: string[] = [];
      if (hf.differentFirst) hfParts.push(` differentFirst="1"`);
      if (hf.differentOddEven) hfParts.push(` differentOddEven="1"`);
      const hfContent: string[] = [];
      if (hf.oddHeader) hfContent.push(`<oddHeader>${escapeXml(hf.oddHeader)}</oddHeader>`);
      if (hf.oddFooter) hfContent.push(`<oddFooter>${escapeXml(hf.oddFooter)}</oddFooter>`);
      p.push(`<headerFooter${hfParts.join("")}>${hfContent.join("")}</headerFooter>`);
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
      if (parseOnOff(attr(sheetPr, "published"))) result.published = true;
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

    // sheetProtection
    const sheetProtectionEl = findChild(el, "sheetProtection");
    if (sheetProtectionEl) {
      const sp: Partial<ChartsheetProtectionOptions> = {};
      if (parseOnOff(attr(sheetProtectionEl, "content"))) sp.content = true;
      if (parseOnOff(attr(sheetProtectionEl, "objects"))) sp.objects = true;
      result.sheetProtection = sp;
    }

    // pageMargins
    const pageMarginsEl = findChild(el, "pageMargins");
    if (pageMarginsEl) {
      const pm: Partial<ChartsheetPageMargins> = {};
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

    // pageSetup
    const pageSetupEl = findChild(el, "pageSetup");
    if (pageSetupEl) {
      const ps: Partial<ChartsheetPageSetup> = {};
      const pz = attrNum(pageSetupEl, "paperSize");
      if (pz !== undefined) ps.paperSize = pz;
      const orient = attr(pageSetupEl, "orientation");
      if (orient) ps.orientation = orient;
      const hdpi = attrNum(pageSetupEl, "horizontalDpi");
      if (hdpi !== undefined) ps.horizontalDpi = hdpi;
      const vdpi = attrNum(pageSetupEl, "verticalDpi");
      if (vdpi !== undefined) ps.verticalDpi = vdpi;
      const copies = attrNum(pageSetupEl, "copies");
      if (copies !== undefined) ps.copies = copies;
      result.pageSetup = ps;
    }

    // headerFooter
    const headerFooterEl = findChild(el, "headerFooter");
    if (headerFooterEl) {
      const hf: Partial<ChartsheetHeaderFooterOptions> = {};
      if (parseOnOff(attr(headerFooterEl, "differentFirst"))) hf.differentFirst = true;
      if (parseOnOff(attr(headerFooterEl, "differentOddEven"))) hf.differentOddEven = true;
      const oh = findChild(headerFooterEl, "oddHeader");
      if (oh) hf.oddHeader = textOf(oh);
      const of2 = findChild(headerFooterEl, "oddFooter");
      if (of2) hf.oddFooter = textOf(of2);
      result.headerFooter = hf;
    }

    return result as ChartsheetDescriptorOptions;
  },
};
