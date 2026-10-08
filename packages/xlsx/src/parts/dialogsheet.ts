/**
 * Dialogsheet types and descriptor for SpreadsheetML documents.
 *
 * A dialogsheet is a legacy Excel 5.0 dialog sheet (no cell data).
 *
 * Reference: OOXML transitional, sml.xsd, CT_Dialogsheet
 *
 * @module
 */

import { convertToInch, parseOnOff } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attrs, attr, attrNum, escapeXml, findChild, stringifyElement } from "@office-open/xml";

import {
  parseHeaderFooterEl,
  parsePageSetupEl,
  parseSheetProtectionEl,
} from "./worksheet/descriptor";
import {
  buildSheetViewAttrs,
  stringifyHeaderFooterXml,
  stringifyPageSetupXml,
  stringifyPrintOptionsXml,
  stringifySheetProtectionXml,
} from "./worksheet/stringify";
import type {
  HeaderFooterOptions,
  PageMarginsOptions,
  PageSetupOptions,
  PrintOptions,
  SheetProtectionOptions,
  SheetViewOptions,
  SheetFormatPropertiesOptions,
} from "./worksheet/types";

// ── Types ──

/** A legacy dialog sheet (xl/dialogsheets/sheetN.xml) — form-dialog controls. */
export interface DialogsheetOptions {
  /** Source package path; round-trip only. */
  sourcePath?: string;
  /** Sheet name */
  name?: string;
  /** Workbook sheet id (CT_Sheet `@sheetId`) — unique but not necessarily sequential. */
  sheetId?: number;
  /** Visibility (CT_Sheet `@state`) */
  state?: "visible" | "hidden" | "veryHidden";
  /** Tab color (hex ARGB) */
  tabColor?: string;
  /** Published to a server (CT_SheetPr `@published`, XSD default true — only false is emitted) */
  published?: boolean;
  /** VBA code name (CT_SheetPr `@codeName`) */
  codeName?: string;
  /** Sheet view (CT_SheetView) */
  sheetView?: SheetViewOptions;
  /** Sheet format properties (CT_SheetFormatPr) */
  sheetFormat?: SheetFormatPropertiesOptions;
  pageMargins?: PageMarginsOptions;
  pageSetup?: PageSetupOptions;
  sheetProtection?: SheetProtectionOptions;
  /** Print options (CT_PrintOptions) */
  printOptions?: PrintOptions;
  /** Header/footer (CT_HeaderFooter) */
  headerFooter?: HeaderFooterOptions;
  /**
   * Relationship id of the dialog form drawing (CT_Dialogsheet `drawing`).
   * Round-trip only: the referenced drawing part is not re-emitted, so the id
   * is not resolvable in a freshly generated workbook.
   */
  drawingRId?: string;
  /**
   * Relationship id of the legacy VML drawing (CT_Dialogsheet `legacyDrawing`).
   * Round-trip only: the referenced VML part is not re-emitted.
   */
  legacyDrawingRId?: string;
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

    // sheetViews (optional) — CT_SheetViews, shared with worksheet
    if (opts.sheetView) {
      p.push(`<sheetViews><sheetView${buildSheetViewAttrs(opts.sheetView)}/></sheetViews>`);
    }

    // sheetFormatPr (optional) — CT_SheetFormatPr, shared with worksheet
    if (opts.sheetFormat) {
      const sfp = opts.sheetFormat;
      const sfpAttrs: Record<string, string | number | boolean | undefined> = {};
      if (sfp.baseColWidth !== undefined) sfpAttrs.baseColWidth = sfp.baseColWidth;
      if (sfp.defaultColWidth !== undefined) sfpAttrs.defaultColWidth = sfp.defaultColWidth;
      if (sfp.defaultRowHeight !== undefined) sfpAttrs.defaultRowHeight = sfp.defaultRowHeight;
      if (sfp.customHeight !== undefined) sfpAttrs.customHeight = sfp.customHeight ? 1 : 0;
      if (sfp.zeroHeight !== undefined) sfpAttrs.zeroHeight = sfp.zeroHeight ? 1 : 0;
      if (sfp.thickTop !== undefined) sfpAttrs.thickTop = sfp.thickTop ? 1 : 0;
      if (sfp.thickBottom !== undefined) sfpAttrs.thickBottom = sfp.thickBottom ? 1 : 0;
      if (sfp.outlineLevelRow !== undefined) sfpAttrs.outlineLevelRow = sfp.outlineLevelRow;
      if (sfp.outlineLevelCol !== undefined) sfpAttrs.outlineLevelCol = sfp.outlineLevelCol;
      p.push(`<sheetFormatPr${attrs(sfpAttrs)}/>`);
    }

    // sheetProtection (optional) — CT_SheetProtection, shared with worksheet
    if (opts.sheetProtection) {
      p.push(stringifySheetProtectionXml(opts.sheetProtection));
    }

    // printOptions (optional) — CT_PrintOptions, shared with worksheet
    if (opts.printOptions) {
      p.push(stringifyPrintOptionsXml(opts.printOptions));
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

    // headerFooter (optional) — CT_HeaderFooter, shared with worksheet
    if (opts.headerFooter) {
      const hfXml = stringifyHeaderFooterXml(opts.headerFooter);
      if (hfXml) p.push(hfXml);
    }

    // drawing / legacyDrawing (optional) — r:id passthrough, round-trip only
    if (opts.drawingRId) {
      p.push(`<drawing r:id="${escapeXml(opts.drawingRId)}"/>`);
    }
    if (opts.legacyDrawingRId) {
      p.push(`<legacyDrawing r:id="${escapeXml(opts.legacyDrawingRId)}"/>`);
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

    // sheetViews — CT_SheetViews, shared with worksheet
    const sheetViewsEl = findChild(el, "sheetViews");
    if (sheetViewsEl) {
      const svEl = findChild(sheetViewsEl, "sheetView");
      if (svEl) {
        const sv: SheetViewOptions = {};
        const sheetViewFlag = (name: string): boolean | undefined =>
          attr(svEl, name) !== undefined ? parseOnOff(attr(svEl, name)) : undefined;
        const showGridLines = sheetViewFlag("showGridLines");
        if (showGridLines !== undefined) sv.showGridLines = showGridLines;
        const showRowColHeaders = sheetViewFlag("showRowColHeaders");
        if (showRowColHeaders !== undefined) sv.showRowColHeaders = showRowColHeaders;
        const showZeros = sheetViewFlag("showZeros");
        if (showZeros !== undefined) sv.showZeros = showZeros;
        const zs = attrNum(svEl, "zoomScale");
        if (zs !== undefined) sv.zoomScale = zs;
        if (attr(svEl, "tabSelected") !== undefined)
          sv.tabSelected = parseOnOff(attr(svEl, "tabSelected")) ?? true;
        const rightToLeft = sheetViewFlag("rightToLeft");
        if (rightToLeft !== undefined) sv.rightToLeft = rightToLeft;
        const windowProtection = sheetViewFlag("windowProtection");
        if (windowProtection !== undefined) sv.windowProtection = windowProtection;
        const showFormulas = sheetViewFlag("showFormulas");
        if (showFormulas !== undefined) sv.showFormulas = showFormulas;
        const showRuler = sheetViewFlag("showRuler");
        if (showRuler !== undefined) sv.showRuler = showRuler;
        const showOutlineSymbols = sheetViewFlag("showOutlineSymbols");
        if (showOutlineSymbols !== undefined) sv.showOutlineSymbols = showOutlineSymbols;
        const defaultGridColor = sheetViewFlag("defaultGridColor");
        if (defaultGridColor !== undefined) sv.defaultGridColor = defaultGridColor;
        const showWhiteSpace = sheetViewFlag("showWhiteSpace");
        if (showWhiteSpace !== undefined) sv.showWhiteSpace = showWhiteSpace;
        const viewVal = attr(svEl, "view");
        if (viewVal) sv.view = viewVal as SheetViewOptions["view"];
        const topLeftCell = attr(svEl, "topLeftCell");
        if (topLeftCell) sv.topLeftCell = topLeftCell;
        const colorId = attrNum(svEl, "colorId");
        if (colorId !== undefined) sv.colorId = colorId;
        const zsn = attrNum(svEl, "zoomScaleNormal");
        if (zsn !== undefined) sv.zoomScaleNormal = zsn;
        const zssl = attrNum(svEl, "zoomScaleSheetLayoutView");
        if (zssl !== undefined) sv.zoomScaleSheetLayoutView = zssl;
        const zspl = attrNum(svEl, "zoomScalePageLayoutView");
        if (zspl !== undefined) sv.zoomScalePageLayoutView = zspl;
        result.sheetView = sv;
      }
    }

    // sheetFormatPr — CT_SheetFormatPr, shared with worksheet
    const sfpEl = findChild(el, "sheetFormatPr");
    if (sfpEl) {
      const sfp: SheetFormatPropertiesOptions = {};
      const bcw = attrNum(sfpEl, "baseColWidth");
      if (bcw !== undefined) sfp.baseColWidth = bcw;
      const dcw = attrNum(sfpEl, "defaultColWidth");
      if (dcw !== undefined) sfp.defaultColWidth = dcw;
      const drh = attrNum(sfpEl, "defaultRowHeight");
      if (drh !== undefined) sfp.defaultRowHeight = drh;
      const ch = attr(sfpEl, "customHeight");
      if (ch !== undefined) sfp.customHeight = parseOnOff(ch) ?? false;
      const zh = attr(sfpEl, "zeroHeight");
      if (zh !== undefined) sfp.zeroHeight = parseOnOff(zh) ?? false;
      const tt = attr(sfpEl, "thickTop");
      if (tt !== undefined) sfp.thickTop = parseOnOff(tt) ?? false;
      const tb = attr(sfpEl, "thickBottom");
      if (tb !== undefined) sfp.thickBottom = parseOnOff(tb) ?? false;
      const olr = attrNum(sfpEl, "outlineLevelRow");
      if (olr !== undefined) sfp.outlineLevelRow = olr;
      const olc = attrNum(sfpEl, "outlineLevelCol");
      if (olc !== undefined) sfp.outlineLevelCol = olc;
      result.sheetFormat = sfp;
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

    // printOptions — CT_PrintOptions, shared with worksheet
    const poEl = findChild(el, "printOptions");
    if (poEl) {
      const po: PrintOptions = {};
      if (parseOnOff(attr(poEl, "horizontalCentered"))) po.horizontalCentered = true;
      if (parseOnOff(attr(poEl, "verticalCentered"))) po.verticalCentered = true;
      if (parseOnOff(attr(poEl, "headings"))) po.headings = true;
      if (parseOnOff(attr(poEl, "gridLines"))) po.gridLines = true;
      if (String(attr(poEl, "gridLinesSet")) === "0") po.gridLinesSet = false;
      result.printOptions = po;
    }

    // headerFooter — CT_HeaderFooter, shared with worksheet
    const hfEl = findChild(el, "headerFooter");
    if (hfEl) {
      result.headerFooter = parseHeaderFooterEl(hfEl, "xl/dialogsheets/dialogsheet.xml");
    }

    // drawing / legacyDrawing — r:id passthrough
    const drEl = findChild(el, "drawing");
    if (drEl) {
      const rId = drEl.attributes?.["r:id"] as string | undefined;
      if (rId) result.drawingRId = rId;
    }
    const ldEl = findChild(el, "legacyDrawing");
    if (ldEl) {
      const rId = ldEl.attributes?.["r:id"] as string | undefined;
      if (rId) result.legacyDrawingRId = rId;
    }

    // extLst — preserved verbatim
    const extEl = findChild(el, "extLst");
    if (extEl) result.extLst = stringifyElement(extEl);

    return result as DialogsheetOptions;
  },
};
