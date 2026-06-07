/**
 * Worksheet descriptor for XLSX — generates xl/worksheets/sheet{n}.xml.
 *
 * Follows the PPTX slideDesc pattern: a large CustomDescriptor with
 * manual `parts.push/join` XML building in stringify(), and parse()
 * for the read path. The cell hot path stays inline for performance.
 *
 * @module
 */

import { derivePasswordHash } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attrs, escapeXml, selfCloseElement } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { findChild, attr, attrNum, textOf } from "@office-open/xml";

import type { SharedStrings } from "../../file/shared-strings";
import { buildRstXml } from "../../file/shared-strings";
import type { Styles } from "../../file/styles";
import type {
  WorksheetOptions,
  WorksheetContext,
  CellOptions,
  FormulaOptions,
  SheetViewOptions,
  SelectionOptions,
  PivotSelectionOptions,
  CfvoOptions,
} from "../../file/worksheet";
import { FormulaType } from "../../file/worksheet";

// ── Descriptor ──

export const worksheetDesc: CustomDescriptor<WorksheetOptions> = {
  kind: "custom",

  /**
   * NOT intended for direct use by the compiler.
   * The compiler calls `stringifyWorksheet(opts, ctx)` instead, which has
   * access to the SharedStrings and Styles accumulators.
   * This method exists to satisfy the CustomDescriptor interface for the read path.
   */
  stringify(_opts, _ctx) {
    throw new Error(
      "Use stringifyWorksheet(opts, ctx) for the write path. worksheetDesc.stringify() is not supported.",
    );
  },

  parse(el, _ctx) {
    const result: Record<string, unknown> = {};

    // Sheet properties
    const sheetPrEl = findChild(el, "sheetPr");
    if (sheetPrEl) {
      const sp: Record<string, unknown> = {};
      if (attr(sheetPrEl, "syncHorizontal") === "1") sp.syncHorizontal = true;
      if (attr(sheetPrEl, "syncVertical") === "1") sp.syncVertical = true;
      if (attr(sheetPrEl, "syncRef")) sp.syncRef = attr(sheetPrEl, "syncRef");
      if (attr(sheetPrEl, "transitionEvaluation") === "1") sp.transitionEvaluation = true;
      if (attr(sheetPrEl, "transitionEntry") === "1") sp.transitionEntry = true;
      if (attr(sheetPrEl, "published") === "1") sp.published = true;
      if (attr(sheetPrEl, "filterMode") === "1") sp.filterMode = true;
      if (attr(sheetPrEl, "enableFormatConditionsCalculation") === "1")
        sp.enableFormatConditionsCalculation = true;

      const outlinePr = findChild(sheetPrEl, "outlinePr");
      if (outlinePr) {
        if (attr(outlinePr, "applyStyles") === "1") sp.outlineApplyStyles = true;
        if (attr(outlinePr, "showOutlineSymbols") === "0") sp.outlineShowSymbols = false;
      }
      if (Object.keys(sp).length > 0) result.sheetPr = sp;

      // Tab color
      const tabColorEl = findChild(sheetPrEl, "tabColor");
      if (tabColorEl) {
        const tc: Record<string, unknown> = {};
        if (attr(tabColorEl, "rgb")) tc.rgb = attr(tabColorEl, "rgb");
        if (attrNum(tabColorEl, "theme") !== undefined) tc.theme = attrNum(tabColorEl, "theme");
        if (attrNum(tabColorEl, "tint") !== undefined) tc.tint = attrNum(tabColorEl, "tint");
        if (attrNum(tabColorEl, "indexed") !== undefined)
          tc.indexed = attrNum(tabColorEl, "indexed");
        result.tabColor = tc;
      }
    }

    // Sheet views
    const sheetViewsEl = findChild(el, "sheetViews");
    if (sheetViewsEl) {
      const svEl = findChild(sheetViewsEl, "sheetView");
      if (svEl) {
        const sv: Record<string, unknown> = {};
        if (attr(svEl, "showGridLines") === "0") sv.showGridLines = false;
        if (attr(svEl, "showRowColHeaders") === "0") sv.showRowColHeaders = false;
        if (attr(svEl, "showZeros") === "0") sv.showZeros = false;
        const zs = attrNum(svEl, "zoomScale");
        if (zs !== undefined) sv.zoomScale = zs;
        if (attr(svEl, "tabSelected") !== undefined)
          sv.tabSelected = attr(svEl, "tabSelected") !== "0";
        if (attr(svEl, "rightToLeft") === "1") sv.rightToLeft = true;
        if (attr(svEl, "view")) sv.view = attr(svEl, "view");
        result.sheetView = sv;

        // Freeze pane
        const paneEl = findChild(svEl, "pane");
        if (paneEl && attr(paneEl, "state") === "frozen") {
          const fp: Record<string, unknown> = {};
          const ys = attrNum(paneEl, "ySplit");
          if (ys && ys > 0) fp.row = ys;
          const xs = attrNum(paneEl, "xSplit");
          if (xs && xs > 0) fp.col = xs;
          if (Object.keys(fp).length > 0) result.freezePanes = fp;
        }
      }
    }

    // Sheet format properties
    const sfpEl = findChild(el, "sheetFormatPr");
    if (sfpEl) {
      const sfp: Record<string, unknown> = {};
      const bcw = attrNum(sfpEl, "baseColWidth");
      if (bcw !== undefined) sfp.baseColWidth = bcw;
      const dcw = attrNum(sfpEl, "defaultColWidth");
      if (dcw !== undefined) sfp.defaultColWidth = dcw;
      const drh = attrNum(sfpEl, "defaultRowHeight");
      if (drh !== undefined) sfp.defaultRowHeight = drh;
      if (attr(sfpEl, "zeroHeight") === "1") sfp.zeroHeight = true;
      if (attr(sfpEl, "thickTop") === "1") sfp.thickTop = true;
      if (attr(sfpEl, "thickBottom") === "1") sfp.thickBottom = true;
      const olr = attrNum(sfpEl, "outlineLevelRow");
      if (olr !== undefined) sfp.outlineLevelRow = olr;
      const olc = attrNum(sfpEl, "outlineLevelCol");
      if (olc !== undefined) sfp.outlineLevelCol = olc;
      result.sheetFormatPr = sfp;
    }

    // Columns
    const colsEl = findChild(el, "cols");
    if (colsEl) {
      const columns: Record<string, unknown>[] = [];
      for (const colEl of colsEl.elements ?? []) {
        if (colEl.name !== "col") continue;
        const col: Record<string, unknown> = {};
        col.min = attrNum(colEl, "min") ?? 0;
        col.max = attrNum(colEl, "max") ?? 0;
        const w = attrNum(colEl, "width");
        if (w !== undefined) col.width = w;
        if (attr(colEl, "hidden") === "1") col.hidden = true;
        if (attr(colEl, "customWidth") === "1") col.customWidth = true;
        const ol = attrNum(colEl, "outlineLevel");
        if (ol !== undefined) col.outlineLevel = ol;
        if (attr(colEl, "collapsed") === "1") col.collapsed = true;
        if (attr(colEl, "bestFit") === "1") col.bestFit = true;
        if (attr(colEl, "phonetic") === "1") col.phonetic = true;
        columns.push(col);
      }
      if (columns.length > 0) result.columns = columns;
    }

    // Sheet protection
    const protEl = findChild(el, "sheetProtection");
    if (protEl?.attributes) {
      const prot: Record<string, unknown> = {};
      if (attr(protEl, "password")) prot.password = attr(protEl, "password");
      if (attr(protEl, "algorithmName")) prot.algorithmName = attr(protEl, "algorithmName");
      if (attr(protEl, "hashValue")) prot.hashValue = attr(protEl, "hashValue");
      if (attr(protEl, "saltValue")) prot.saltValue = attr(protEl, "saltValue");
      if (attrNum(protEl, "spinCount") !== undefined) prot.spinCount = attrNum(protEl, "spinCount");
      if (attr(protEl, "sheet") === "1") prot.sheet = true;
      if (attr(protEl, "objects") === "1") prot.objects = true;
      if (attr(protEl, "scenarios") === "1") prot.scenarios = true;
      if (attr(protEl, "formatCells") === "0") prot.formatCells = false;
      if (attr(protEl, "formatColumns") === "0") prot.formatColumns = false;
      if (attr(protEl, "formatRows") === "0") prot.formatRows = false;
      if (attr(protEl, "insertColumns") === "0") prot.insertColumns = false;
      if (attr(protEl, "insertRows") === "0") prot.insertRows = false;
      if (attr(protEl, "insertHyperlinks") === "0") prot.insertHyperlinks = false;
      if (attr(protEl, "deleteColumns") === "0") prot.deleteColumns = false;
      if (attr(protEl, "deleteRows") === "0") prot.deleteRows = false;
      if (attr(protEl, "selectLockedCells") === "1") prot.selectLockedCells = true;
      if (attr(protEl, "sort") === "0") prot.sort = false;
      if (attr(protEl, "autoFilter") === "0") prot.autoFilter = false;
      if (attr(protEl, "pivotTables") === "0") prot.pivotTables = false;
      if (attr(protEl, "selectUnlockedCells") === "1") prot.selectUnlockedCells = true;
      result.protection = prot;
    }

    // Protected ranges
    const prEl = findChild(el, "protectedRanges");
    if (prEl) {
      const ranges: Record<string, unknown>[] = [];
      for (const rEl of prEl.elements ?? []) {
        if (rEl.name !== "protectedRange") continue;
        const r: Record<string, unknown> = {};
        r.sqref = attr(rEl, "sqref") ?? "";
        r.name = attr(rEl, "name") ?? "";
        if (attr(rEl, "password")) r.password = attr(rEl, "password");
        if (attr(rEl, "algorithmName")) r.algorithmName = attr(rEl, "algorithmName");
        if (attr(rEl, "hashValue")) r.hashValue = attr(rEl, "hashValue");
        if (attr(rEl, "saltValue")) r.saltValue = attr(rEl, "saltValue");
        if (attrNum(rEl, "spinCount") !== undefined) r.spinCount = attrNum(rEl, "spinCount");
        const sdEl = findChild(rEl, "securityDescriptor");
        if (sdEl) r.securityDescriptor = textOf(sdEl);
        ranges.push(r);
      }
      if (ranges.length > 0) result.protectedRanges = ranges;
    }

    // Auto filter
    const afEl = findChild(el, "autoFilter");
    if (afEl) {
      result.autoFilter = attr(afEl, "ref") ?? "";
    }

    // Merge cells
    const mcEl = findChild(el, "mergeCells");
    if (mcEl) {
      const merges: Record<string, unknown>[] = [];
      for (const mEl of mcEl.elements ?? []) {
        if (mEl.name !== "mergeCell") continue;
        const ref = attr(mEl, "ref") ?? "";
        const parts = ref.split(":");
        if (parts.length === 2) {
          const from = parseCellRef(parts[0]);
          const to = parseCellRef(parts[1]);
          if (from && to) merges.push({ from, to });
        }
      }
      if (merges.length > 0) result.mergeCells = merges;
    }

    // Conditional formatting
    const cfEls = el.elements?.filter((e) => e.name === "conditionalFormatting") ?? [];
    if (cfEls.length > 0) {
      const cfs: Record<string, unknown>[] = [];
      for (const cfEl of cfEls) {
        const sqref = attr(cfEl, "sqref") ?? "";
        const rules: Record<string, unknown>[] = [];
        for (const ruleEl of cfEl.elements ?? []) {
          if (ruleEl.name !== "cfRule") continue;
          const rule: Record<string, unknown> = {};
          rule.type = attr(ruleEl, "type");
          rule.priority = attrNum(ruleEl, "priority") ?? 1;
          if (attr(ruleEl, "operator")) rule.operator = attr(ruleEl, "operator");
          const dxfId = attrNum(ruleEl, "dxfId");
          if (dxfId !== undefined) rule.dxfId = dxfId;
          if (attr(ruleEl, "stopIfTrue") === "1") rule.stopIfTrue = true;
          if (attr(ruleEl, "timePeriod")) rule.timePeriod = attr(ruleEl, "timePeriod");
          const rank = attrNum(ruleEl, "rank");
          if (rank !== undefined) rule.rank = rank;
          if (attr(ruleEl, "equalAverage") === "1") rule.equalAverage = true;

          // Color scale
          const csEl = findChild(ruleEl, "colorScale");
          if (csEl) {
            const cfvo: Record<string, unknown>[] = [];
            const colors: string[] = [];
            for (const child of csEl.elements ?? []) {
              if (child.name === "cfvo") cfvo.push(parseCfvo(child));
              if (child.name === "color") {
                const rgb = attr(child, "rgb");
                if (rgb) colors.push(rgb.length === 8 ? rgb.slice(2) : rgb);
              }
            }
            rule.colorScale = { cfvo, colors };
          }

          // Data bar
          const dbEl = findChild(ruleEl, "dataBar");
          if (dbEl) {
            const cfvo: Record<string, unknown>[] = [];
            let color = "";
            for (const child of dbEl.elements ?? []) {
              if (child.name === "cfvo") cfvo.push(parseCfvo(child));
              if (child.name === "color") {
                const rgb = attr(child, "rgb");
                if (rgb) color = rgb.length === 8 ? rgb.slice(2) : rgb;
              }
            }
            rule.dataBar = { cfvo: cfvo as [any, any], color };
          }

          // Icon set
          const isEl = findChild(ruleEl, "iconSet");
          if (isEl) {
            const cfvo: Record<string, unknown>[] = [];
            for (const child of isEl.elements ?? []) {
              if (child.name === "cfvo") cfvo.push(parseCfvo(child));
            }
            const iconSet: Record<string, unknown> = { cfvo };
            if (attr(isEl, "iconSet")) iconSet.iconSet = attr(isEl, "iconSet");
            if (attr(isEl, "showValue") === "0") iconSet.showValue = false;
            if (attr(isEl, "percent") === "0") iconSet.percent = false;
            if (attr(isEl, "reverse") === "1") iconSet.reverse = true;
            rule.iconSet = iconSet;
          }

          // Formulas
          const formulas: string[] = [];
          for (const child of ruleEl.elements ?? []) {
            if (child.name === "formula") formulas.push(textOf(child) ?? "");
          }
          if (formulas.length > 0) rule.formulas = formulas;

          rules.push(rule);
        }
        cfs.push({ sqref, rules });
      }
      result.conditionalFormats = cfs;
    }

    // Data validations
    const dvEl = findChild(el, "dataValidations");
    if (dvEl) {
      const dvs: Record<string, unknown>[] = [];
      for (const dEl of dvEl.elements ?? []) {
        if (dEl.name !== "dataValidation") continue;
        const dv: Record<string, unknown> = {};
        dv.sqref = attr(dEl, "sqref") ?? "";
        if (attr(dEl, "type")) dv.type = attr(dEl, "type");
        if (attr(dEl, "operator")) dv.operator = attr(dEl, "operator");
        if (attr(dEl, "allowBlank") === "1") dv.allowBlank = true;
        if (attr(dEl, "showErrorMessage") === "1") dv.showErrorMessage = true;
        if (attr(dEl, "showInputMessage") === "1") dv.showInputMessage = true;
        if (attr(dEl, "errorTitle")) dv.errorTitle = attr(dEl, "errorTitle");
        if (attr(dEl, "error")) dv.error = attr(dEl, "error");
        if (attr(dEl, "promptTitle")) dv.promptTitle = attr(dEl, "promptTitle");
        if (attr(dEl, "prompt")) dv.prompt = attr(dEl, "prompt");
        if (attr(dEl, "errorStyle")) dv.errorStyle = attr(dEl, "errorStyle");
        if (attr(dEl, "imeMode")) dv.imeMode = attr(dEl, "imeMode");
        if (attr(dEl, "showDropDown") === "1") dv.showDropDown = true;

        const f1El = findChild(dEl, "formula1");
        if (f1El) dv.formula1 = textOf(f1El);
        const f2El = findChild(dEl, "formula2");
        if (f2El) dv.formula2 = textOf(f2El);

        dvs.push(dv);
      }
      result.dataValidations = dvs;
    }

    // Hyperlinks
    const hlEl = findChild(el, "hyperlinks");
    if (hlEl) {
      const hyperlinks: Record<string, unknown>[] = [];
      for (const hEl of hlEl.elements ?? []) {
        if (hEl.name !== "hyperlink") continue;
        const hl: Record<string, unknown> = {};
        hl.cell = attr(hEl, "ref") ?? "";
        const rId = hEl.attributes?.["r:id"] as string | undefined;
        const location = attr(hEl, "location");
        if (rId) hl.target = { type: "external", url: rId };
        else if (location) hl.target = { type: "internal", location };
        if (attr(hEl, "tooltip")) hl.tooltip = attr(hEl, "tooltip");
        if (attr(hEl, "display")) hl.display = attr(hEl, "display");
        hyperlinks.push(hl);
      }
      result.hyperlinks = hyperlinks;
    }

    // Print options
    const poEl = findChild(el, "printOptions");
    if (poEl) {
      const po: Record<string, unknown> = {};
      if (attr(poEl, "horizontalCentered") === "1") po.horizontalCentered = true;
      if (attr(poEl, "verticalCentered") === "1") po.verticalCentered = true;
      if (attr(poEl, "headings") === "1") po.headings = true;
      if (attr(poEl, "gridLines") === "1") po.gridLines = true;
      if (attr(poEl, "gridLinesSet") === "0") po.gridLinesSet = false;
      result.printOptions = po;
    }

    // Page setup
    const psEl = findChild(el, "pageSetup");
    if (psEl) {
      const ps: Record<string, unknown> = {};
      const pz = attrNum(psEl, "paperSize");
      if (pz !== undefined) ps.paperSize = pz;
      if (attr(psEl, "orientation")) ps.orientation = attr(psEl, "orientation");
      const sc = attrNum(psEl, "scale");
      if (sc !== undefined) ps.scale = sc;
      const ftw = attrNum(psEl, "fitToWidth");
      if (ftw !== undefined) ps.fitToWidth = ftw;
      const fth = attrNum(psEl, "fitToHeight");
      if (fth !== undefined) ps.fitToHeight = fth;
      if (attr(psEl, "pageOrder")) ps.pageOrder = attr(psEl, "pageOrder");
      if (attr(psEl, "useFirstPageNumber") === "1") ps.useFirstPageNumber = true;
      const fpn = attrNum(psEl, "firstPageNumber");
      if (fpn !== undefined) ps.firstPageNumber = fpn;
      result.pageSetup = ps;
    }

    // Header/footer
    const hfEl = findChild(el, "headerFooter");
    if (hfEl) {
      const hf: Record<string, unknown> = {};
      if (attr(hfEl, "differentOddEven") === "1") hf.differentOddEven = true;
      if (attr(hfEl, "differentFirst") === "1") hf.differentFirst = true;
      if (attr(hfEl, "scaleWithDoc") === "0") hf.scaleWithDoc = false;
      if (attr(hfEl, "alignWithMargins") === "0") hf.alignWithMargins = false;
      const oh = findChild(hfEl, "oddHeader");
      if (oh) hf.oddHeader = textOf(oh);
      const of2 = findChild(hfEl, "oddFooter");
      if (of2) hf.oddFooter = textOf(of2);
      const eh = findChild(hfEl, "evenHeader");
      if (eh) hf.evenHeader = textOf(eh);
      const ef = findChild(hfEl, "evenFooter");
      if (ef) hf.evenFooter = textOf(ef);
      const fh = findChild(hfEl, "firstHeader");
      if (fh) hf.firstHeader = textOf(fh);
      const ff = findChild(hfEl, "firstFooter");
      if (ff) hf.firstFooter = textOf(ff);
      result.headerFooter = hf;
    }

    // Ignored errors
    const ieEl = findChild(el, "ignoredErrors");
    if (ieEl) {
      const errors: Record<string, unknown>[] = [];
      for (const eEl of ieEl.elements ?? []) {
        if (eEl.name !== "ignoredError") continue;
        const ie: Record<string, unknown> = {};
        ie.sqref = attr(eEl, "sqref") ?? "";
        if (attr(eEl, "evalError") === "1") ie.evalError = true;
        if (attr(eEl, "twoDigitTextYear") === "1") ie.twoDigitTextYear = true;
        if (attr(eEl, "numberStoredAsText") === "1") ie.numberStoredAsText = true;
        if (attr(eEl, "formula") === "1") ie.formula = true;
        if (attr(eEl, "formulaRange") === "1") ie.formulaRange = true;
        if (attr(eEl, "unlockedFormula") === "1") ie.unlockedFormula = true;
        if (attr(eEl, "emptyCellReference") === "1") ie.emptyCellReference = true;
        if (attr(eEl, "listDataValidation") === "1") ie.listDataValidation = true;
        if (attr(eEl, "calculatedColumn") === "1") ie.calculatedColumn = true;
        errors.push(ie);
      }
      result.ignoredErrors = errors;
    }

    // Phonetic properties
    const ppEl = findChild(el, "phoneticPr");
    if (ppEl) {
      const pp: Record<string, unknown> = {};
      pp.fontId = attrNum(ppEl, "fontId") ?? 0;
      if (attr(ppEl, "type")) pp.type = attr(ppEl, "type");
      if (attr(ppEl, "alignment")) pp.alignment = attr(ppEl, "alignment");
      result.phoneticPr = pp;
    }

    // Sheet calc properties
    const scEl = findChild(el, "sheetCalcPr");
    if (scEl) {
      const sc: Record<string, unknown> = {};
      if (attr(scEl, "fullCalcOnLoad") === "1") sc.fullCalcOnLoad = true;
      result.sheetCalcPr = sc;
    }

    return result as Record<string, unknown>;
  },
};

// ── Stringify implementation ──

/**
 * Build the complete worksheet XML string.
 *
 * Zero-allocation fast path: directly concatenates XML string,
 * bypassing the IXmlableObject intermediate tree entirely.
 */
export function stringifyWorksheet(opts: WorksheetOptions, ctx: WorksheetContext): string {
  const sharedStrings = ctx.sharedStrings;
  const styles = ctx.styles;

  const rows = opts.rows ?? [];
  const columns = opts.columns ?? [];
  const mergeCells = opts.mergeCells ?? [];
  const protectedRanges = opts.protectedRanges ?? [];
  const ignoredErrors = opts.ignoredErrors ?? [];
  const rowBreaks = opts.rowBreaks ?? [];
  const colBreaks = opts.colBreaks ?? [];
  const customSheetViews = opts.customSheetViews ?? [];
  const cellWatches = opts.cellWatches ?? [];
  const controls = opts.controls ?? [];
  const customProperties = opts.customProperties ?? [];
  const oleObjects = opts.oleObjects ?? [];
  const webPublishItems = opts.webPublishItems ?? [];

  const p: string[] = [
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
      ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
      ' mc:Ignorable="x14ac xr xr2 xr3"' +
      ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"' +
      ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"' +
      ' xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"' +
      ' xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3">',
  ];

  // Sheet properties (tabColor, outlinePr go here)
  const hasTabColor = !!opts.tabColor;
  const hasOutline = columns.some((c) => c.outlineLevel !== undefined);
  const sp = opts.sheetPr;
  const hasSheetPrAttrs =
    sp &&
    (sp.syncHorizontal ||
      sp.syncVertical ||
      sp.syncRef ||
      sp.transitionEvaluation ||
      sp.transitionEntry ||
      sp.published ||
      sp.filterMode ||
      sp.enableFormatConditionsCalculation);
  if (hasTabColor || hasOutline || hasSheetPrAttrs) {
    const prParts: string[] = [];
    const prAttrs: Record<string, string | number | boolean | undefined> = {};
    if (sp?.syncHorizontal) prAttrs.syncHorizontal = 1;
    if (sp?.syncVertical) prAttrs.syncVertical = 1;
    if (sp?.syncRef) prAttrs.syncRef = sp.syncRef;
    if (sp?.transitionEvaluation) prAttrs.transitionEvaluation = 1;
    if (sp?.transitionEntry) prAttrs.transitionEntry = 1;
    if (sp?.published) prAttrs.published = 1;
    if (sp?.filterMode) prAttrs.filterMode = 1;
    if (sp?.enableFormatConditionsCalculation) prAttrs.enableFormatConditionsCalculation = 1;
    if (opts.tabColor) {
      const tc = opts.tabColor;
      const tcAttrs: Record<string, string | number | boolean | undefined> = {};
      if (tc.rgb) tcAttrs.rgb = tc.rgb;
      if (tc.theme !== undefined) tcAttrs.theme = tc.theme;
      if (tc.tint !== undefined) tcAttrs.tint = tc.tint;
      if (tc.indexed !== undefined) tcAttrs.indexed = tc.indexed;
      prParts.push(`<tabColor${attrs(tcAttrs)}/>`);
    }
    if (hasOutline) {
      const outAttrs: Record<string, string | number | boolean | undefined> = {
        summaryBelow: 1,
        summaryRight: 1,
      };
      if (sp?.outlineApplyStyles) outAttrs.applyStyles = 1;
      if (sp?.outlineShowSymbols === false) outAttrs.showOutlineSymbols = 0;
      prParts.push(`<outlinePr${attrs(outAttrs)}/>`);
    }
    // pageSetUpPr (inside sheetPr when fitToPage or autoPageBreaks needed)
    if (
      opts.pageSetup?.fitToWidth ||
      opts.pageSetup?.fitToHeight ||
      opts.pageSetup?.autoPageBreaks
    ) {
      const psupAttrs: Record<string, string | number | boolean | undefined> = {};
      if (opts.pageSetup?.fitToWidth || opts.pageSetup?.fitToHeight) psupAttrs.fitToPage = 1;
      if (opts.pageSetup?.autoPageBreaks) psupAttrs.autoPageBreaks = 1;
      prParts.push(`<pageSetUpPr${attrs(psupAttrs)}/>`);
    }
    const prAttrStr = Object.keys(prAttrs).length > 0 ? attrs(prAttrs) : "";
    p.push(`<sheetPr${prAttrStr}>${prParts.join("")}</sheetPr>`);
  }

  // Dimension — defines the used range of the sheet
  const maxRow = rows.length;
  let maxCol = 0;
  for (const row of rows) {
    if (row.cells && row.cells.length > maxCol) maxCol = row.cells.length;
  }
  if (maxRow > 0 && maxCol > 0) {
    const dimRef = `A1:${defaultCellRef(maxRow, maxCol)}`;
    p.push(`<dimension ref="${dimRef}"/>`);
  }

  // Sheet views
  const pivotSelXml = opts.sheetView?.pivotSelections
    ? opts.sheetView.pivotSelections.map((ps) => buildPivotSelectionXml(ps)).join("")
    : "";
  if (opts.freezePanes) {
    const fp = opts.freezePanes;
    const ySplit = fp.row ? fp.row : 0;
    const xSplit = fp.col ? fp.col : 0;
    const topRow = fp.row ? fp.row + 1 : 1;
    const leftCol = fp.col ? fp.col + 1 : 1;
    const topLeftCell = defaultCellRef(topRow, leftCol);
    const activePane =
      ySplit > 0 && xSplit > 0 ? "bottomRight" : ySplit > 0 ? "bottomLeft" : "topRight";
    const svAttrs = buildSheetViewAttrs(opts.sheetView);
    p.push(
      `<sheetViews><sheetView${svAttrs}>`,
      `<pane ySplit="${ySplit}" xSplit="${xSplit}" topLeftCell="${topLeftCell}" activePane="${activePane}" state="frozen"/>`,
      opts.selection ? buildSelectionXml(opts.selection) : "",
      pivotSelXml,
      "</sheetView></sheetViews>",
    );
  } else {
    const svAttrs = buildSheetViewAttrs(opts.sheetView);
    const innerXml = (opts.selection ? buildSelectionXml(opts.selection) : "") + pivotSelXml;
    if (innerXml) {
      p.push(`<sheetViews><sheetView${svAttrs}>${innerXml}</sheetView></sheetViews>`);
    } else {
      p.push(`<sheetViews><sheetView${svAttrs}/></sheetViews>`);
    }
  }

  // Sheet format — default row height
  if (opts.sheetFormatPr) {
    const sfp = opts.sheetFormatPr;
    const sfpAttrs: Record<string, string | number | boolean | undefined> = {};
    if (sfp.baseColWidth !== undefined) sfpAttrs.baseColWidth = sfp.baseColWidth;
    if (sfp.defaultColWidth !== undefined) sfpAttrs.defaultColWidth = sfp.defaultColWidth;
    sfpAttrs.defaultRowHeight = sfp.defaultRowHeight ?? 15;
    if (sfp.zeroHeight) sfpAttrs.zeroHeight = 1;
    if (sfp.thickTop) sfpAttrs.thickTop = 1;
    if (sfp.thickBottom) sfpAttrs.thickBottom = 1;
    if (sfp.outlineLevelRow !== undefined) sfpAttrs.outlineLevelRow = sfp.outlineLevelRow;
    if (sfp.outlineLevelCol !== undefined) sfpAttrs.outlineLevelCol = sfp.outlineLevelCol;
    p.push(`<sheetFormatPr${attrs(sfpAttrs)}/>`);
  } else {
    p.push('<sheetFormatPr defaultRowHeight="15"/>');
  }

  // Column definitions
  if (columns.length > 0) {
    p.push("<cols>");
    for (const col of columns) {
      const colAttrs: Record<string, string | number | boolean | undefined> = {
        min: col.min,
        max: col.max,
      };
      if (col.width !== undefined) {
        colAttrs.width = col.width;
        colAttrs.customWidth = 1;
      }
      if (col.hidden) {
        colAttrs.hidden = 1;
      }
      if (col.outlineLevel !== undefined) {
        colAttrs.outlineLevel = col.outlineLevel;
      }
      if (col.collapsed) {
        colAttrs.collapsed = 1;
      }
      if (col.bestFit) {
        colAttrs.bestFit = 1;
      }
      if (col.phonetic) {
        colAttrs.phonetic = 1;
      }
      p.push(selfCloseElement("col", attrs(colAttrs)));
    }
    p.push("</cols>");
  }

  // Sheet data (rows + cells) — the hot path
  p.push("<sheetData>");
  for (let i = 0; i < rows.length; i++) {
    const rowOpts = rows[i];
    const rowNumber = rowOpts.rowNumber ?? i + 1;
    const rowAttrs: Record<string, string | number | boolean | undefined> = { r: rowNumber };
    if (rowOpts.height !== undefined) {
      rowAttrs.ht = rowOpts.height;
      rowAttrs.customHeight = 1;
    }
    if (rowOpts.hidden) {
      rowAttrs.hidden = 1;
    }
    if (rowOpts.spans) rowAttrs.spans = rowOpts.spans;
    if (rowOpts.customFormat) rowAttrs.customFormat = 1;
    if (rowOpts.thickTop) rowAttrs.thickTop = 1;
    if (rowOpts.thickBot) rowAttrs.thickBot = 1;
    if (rowOpts.ph) rowAttrs.ph = 1;

    if (rowOpts.cells) {
      const rowParts: string[] = [];
      for (let j = 0; j < rowOpts.cells.length; j++) {
        const cell = rowOpts.cells[j];
        const ref = cell.reference ?? defaultCellRef(rowNumber, j + 1);
        const cellStr = buildCellString(ref, cell, sharedStrings, styles);
        if (cellStr) rowParts.push(cellStr);
      }
      p.push(`<row${attrs(rowAttrs)}>`, ...rowParts, "</row>");
    } else {
      p.push(`<row${attrs(rowAttrs)}/>`);
    }
  }
  p.push("</sheetData>");

  // Sheet calc properties (after sheetData per XSD sequence)
  if (opts.sheetCalcPr) {
    const scAttrs: string[] = [];
    if (opts.sheetCalcPr.fullCalcOnLoad) scAttrs.push('fullCalcOnLoad="1"');
    p.push(`<sheetCalcPr${scAttrs.length ? " " + scAttrs.join(" ") : ""}/>`);
  }

  // Row breaks (after sheetCalcPr per XSD sequence)
  if (rowBreaks.length > 0) {
    const brkParts = rowBreaks.map((b) => {
      const bAttrs: Record<string, string | number | boolean | undefined> = { id: b.id };
      if (b.min !== undefined) bAttrs.min = b.min;
      if (b.max !== undefined) bAttrs.max = b.max;
      if (b.manual) bAttrs.man = 1;
      if (b.pivot) bAttrs.pt = 1;
      return `<brk${attrs(bAttrs)}/>`;
    });
    p.push(
      `<rowBreaks count="${rowBreaks.length}" manualBreakCount="${rowBreaks.filter((b) => b.manual).length}">${brkParts.join("")}</rowBreaks>`,
    );
  }

  // Column breaks
  if (colBreaks.length > 0) {
    const brkParts = colBreaks.map((b) => {
      const bAttrs: Record<string, string | number | boolean | undefined> = { id: b.id };
      if (b.min !== undefined) bAttrs.min = b.min;
      if (b.max !== undefined) bAttrs.max = b.max;
      if (b.manual) bAttrs.man = 1;
      if (b.pivot) bAttrs.pt = 1;
      return `<brk${attrs(bAttrs)}/>`;
    });
    p.push(
      `<colBreaks count="${colBreaks.length}" manualBreakCount="${colBreaks.filter((b) => b.manual).length}">${brkParts.join("")}</colBreaks>`,
    );
  }

  // Custom properties (CT_CustomProperties, after colBreaks per XSD sequence)
  if (customProperties.length > 0) {
    const cpParts: string[] = ["<customProperties>"];
    for (const cp of customProperties) {
      cpParts.push(`<customPr name="${escapeXml(cp.name)}" r:id="${escapeXml(cp.rId)}"/>`);
    }
    cpParts.push("</customProperties>");
    p.push(cpParts.join(""));
  }

  // OLE size
  if (opts.oleSize) {
    p.push(`<oleSize ref="${escapeXml(opts.oleSize)}"/>`);
  }

  // Custom sheet views (after oleSize per XSD sequence)
  if (customSheetViews.length > 0) {
    p.push("<customSheetViews>");
    for (const csv of customSheetViews) {
      const csvAttrs: Record<string, string | number | boolean | undefined> = { guid: csv.guid };
      if (csv.scale !== undefined) csvAttrs.scale = csv.scale;
      if (csv.showPageBreaks) csvAttrs.showPageBreaks = 1;
      if (csv.showFormulas) csvAttrs.showFormulas = 1;
      if (csv.showGridLines === false) csvAttrs.showGridLines = 0;
      if (csv.showRowColHeaders === false) csvAttrs.showRowCol = 0;
      if (csv.outlineSymbols === false) csvAttrs.outlineSymbols = 0;
      if (csv.zeroValues === false) csvAttrs.zeroValues = 0;
      if (csv.fitToPage) csvAttrs.fitToPage = 1;
      if (csv.printArea) csvAttrs.printArea = 1;
      if (csv.filter) csvAttrs.filter = 1;
      if (csv.showAutoFilter) csvAttrs.showAutoFilter = 1;
      if (csv.hiddenRows) csvAttrs.hiddenRows = 1;
      if (csv.hiddenColumns) csvAttrs.hiddenColumns = 1;
      if (csv.state && csv.state !== "visible") csvAttrs.state = csv.state;
      if (csv.filterUnique) csvAttrs.filterUnique = 1;
      if (csv.view && csv.view !== "normal") csvAttrs.view = csv.view;
      p.push(`<customSheetView${attrs(csvAttrs)}/>`);
    }
    p.push("</customSheetViews>");
  }

  // Cell watches
  if (cellWatches.length > 0) {
    p.push("<cellWatches>");
    for (const cw of cellWatches) {
      p.push(`<cellWatch r="${escapeXml(cw.r)}"/>`);
    }
    p.push("</cellWatches>");
  }

  // Data consolidation
  if (opts.dataConsolidate) {
    const dc = opts.dataConsolidate;
    const dcAttrs: Record<string, string | number | boolean | undefined> = {};
    if (dc.function && dc.function !== "sum") dcAttrs.function = dc.function;
    if (dc.topLabels) dcAttrs.topLabels = 1;
    if (dc.leftLabels) dcAttrs.leftLabels = 1;
    if (dc.startLabels) dcAttrs.startLabels = 1;
    if (dc.link) dcAttrs.link = 1;
    const refsInner = dc.refs?.map((r) => `<dataRef ref="${escapeXml(r)}"/>`).join("") ?? "";
    const refsXml = refsInner ? `<dataRefs>${refsInner}</dataRefs>` : "";
    if (refsXml || Object.keys(dcAttrs).length > 0) {
      p.push(`<dataConsolidate${attrs(dcAttrs)}>${refsXml}</dataConsolidate>`);
    }
  }

  // Sheet protection (after sheetData, before protectedRanges per XSD sequence)
  if (opts.protection) {
    const prot = opts.protection;
    const protAttrs: Record<string, string | number | boolean | undefined> = {};
    if (prot.password) protAttrs.password = hashPassword(prot.password);
    // Auto-derive modern hash when password provided without explicit hashValue
    let derived: ReturnType<typeof derivePasswordHash> | undefined;
    if (prot.password !== undefined && prot.hashValue === undefined) {
      derived = derivePasswordHash(prot.password);
    }
    protAttrs.algorithmName = prot.algorithmName ?? derived?.algorithmName;
    protAttrs.hashValue = prot.hashValue ?? derived?.hashValue;
    protAttrs.saltValue = prot.saltValue ?? derived?.saltValue;
    if (prot.spinCount !== undefined) protAttrs.spinCount = prot.spinCount;
    else if (derived) protAttrs.spinCount = derived.spinCount;
    if (prot.sheet) protAttrs.sheet = 1;
    if (prot.objects) protAttrs.objects = 1;
    if (prot.scenarios) protAttrs.scenarios = 1;
    if (prot.formatCells === false) protAttrs.formatCells = 0;
    if (prot.formatColumns === false) protAttrs.formatColumns = 0;
    if (prot.formatRows === false) protAttrs.formatRows = 0;
    if (prot.insertColumns === false) protAttrs.insertColumns = 0;
    if (prot.insertRows === false) protAttrs.insertRows = 0;
    if (prot.insertHyperlinks === false) protAttrs.insertHyperlinks = 0;
    if (prot.deleteColumns === false) protAttrs.deleteColumns = 0;
    if (prot.deleteRows === false) protAttrs.deleteRows = 0;
    if (prot.selectLockedCells) protAttrs.selectLockedCells = 1;
    if (prot.sort === false) protAttrs.sort = 0;
    if (prot.autoFilter === false) protAttrs.autoFilter = 0;
    if (prot.pivotTables === false) protAttrs.pivotTables = 0;
    if (prot.selectUnlockedCells) protAttrs.selectUnlockedCells = 1;
    p.push(selfCloseElement("sheetProtection", attrs(protAttrs)));
  }

  // Protected ranges (after sheetProtection per XSD sequence)
  if (protectedRanges.length > 0) {
    const prParts: string[] = ["<protectedRanges>"];
    for (const pr of protectedRanges) {
      const prAttrs: Record<string, string | number | boolean | undefined> = {
        name: pr.name,
        sqref: pr.sqref,
      };
      if (pr.password) prAttrs.password = hashPassword(pr.password);
      // Auto-derive modern hash when password provided without explicit hashValue
      let prDerived: ReturnType<typeof derivePasswordHash> | undefined;
      if (pr.password !== undefined && pr.hashValue === undefined) {
        prDerived = derivePasswordHash(pr.password);
      }
      prAttrs.algorithmName = pr.algorithmName ?? prDerived?.algorithmName;
      prAttrs.hashValue = pr.hashValue ?? prDerived?.hashValue;
      prAttrs.saltValue = pr.saltValue ?? prDerived?.saltValue;
      if (pr.spinCount !== undefined) prAttrs.spinCount = pr.spinCount;
      else if (prDerived) prAttrs.spinCount = prDerived.spinCount;
      const hasSecurityDescriptor = !!pr.securityDescriptor;
      if (hasSecurityDescriptor) {
        prParts.push(
          `<protectedRange${attrs(prAttrs)}><securityDescriptor>${escapeXml(pr.securityDescriptor!)}</securityDescriptor></protectedRange>`,
        );
      } else {
        prParts.push(selfCloseElement("protectedRange", attrs(prAttrs)));
      }
    }
    prParts.push("</protectedRanges>");
    p.push(prParts.join(""));
  }

  // Scenarios (what-if analysis)
  if (opts.scenarios) {
    const scParts: string[] = ["<scenarios"];
    const scAttrs: Record<string, string | number> = {};
    if (opts.scenarios.current !== undefined) scAttrs.current = opts.scenarios.current;
    if (opts.scenarios.show !== undefined) scAttrs.show = opts.scenarios.show;
    scParts[0] = `<scenarios${attrs(scAttrs)}>`;

    for (const scenario of opts.scenarios.scenarios) {
      const sAttrs: Record<string, string | number | boolean | undefined> = {
        name: scenario.name,
      };
      if (scenario.count !== undefined) sAttrs.count = scenario.count;
      if (scenario.user) sAttrs.user = scenario.user;
      if (scenario.comment) sAttrs.comment = scenario.comment;
      if (scenario.hidden) sAttrs.hidden = true;
      if (scenario.locked) sAttrs.locked = true;

      const sParts: string[] = [`<scenario${attrs(sAttrs)}>`];
      for (const cell of scenario.inputCells) {
        const icAttrs: Record<string, string | number | boolean | undefined> = {
          r: cell.r,
          val: String(cell.val),
        };
        if (cell.deleted) icAttrs.deleted = true;
        if (cell.undone) icAttrs.undone = true;
        sParts.push(`<inputCells${attrs(icAttrs)}/>`);
      }
      sParts.push("</scenario>");
      scParts.push(sParts.join(""));
    }
    scParts.push("</scenarios>");
    p.push(scParts.join(""));
  }

  // Auto filter
  if (opts.autoFilter) {
    if (typeof opts.autoFilter === "string") {
      p.push(selfCloseElement("autoFilter", attrs({ ref: opts.autoFilter })));
    } else {
      const af = opts.autoFilter;
      const inner: string[] = [];
      for (const t10 of af.top10 ?? []) {
        const fcAttrs: Record<string, string | number | boolean | undefined> = {
          colId: t10.colId,
        };
        if (t10.hiddenButton) fcAttrs.hiddenButton = 1;
        if (t10.showButton === false) fcAttrs.showButton = 0;
        const t10Attrs: Record<string, string | number | boolean | undefined> = { val: t10.val };
        if (t10.top === false) t10Attrs.top = 0;
        if (t10.percent) t10Attrs.percent = 1;
        if (t10.filterVal !== undefined) t10Attrs.filterVal = t10.filterVal;
        inner.push(`<filterColumn${attrs(fcAttrs)}><top10${attrs(t10Attrs)}/></filterColumn>`);
      }
      for (const cf of af.customFilters ?? []) {
        const fcAttrs: Record<string, string | number | boolean | undefined> = {
          colId: cf.colId,
        };
        if (cf.hiddenButton) fcAttrs.hiddenButton = 1;
        if (cf.showButton === false) fcAttrs.showButton = 0;
        const cfAttrs: Record<string, string | number | boolean | undefined> = {};
        if (cf.and) cfAttrs.and = 1;
        const filters: string[] = [];
        if (cf.val !== undefined) {
          const fAttrs: Record<string, string | number | boolean | undefined> = { val: cf.val };
          if (cf.operator) fAttrs.operator = cf.operator;
          filters.push(selfCloseElement("customFilter", attrs(fAttrs)));
        }
        if (cf.val2 !== undefined) {
          filters.push(selfCloseElement("customFilter", attrs({ val: cf.val2 })));
        }
        if (filters.length > 0) {
          inner.push(
            `<filterColumn${attrs(fcAttrs)}><customFilters${attrs(cfAttrs)}>${filters.join("")}</customFilters></filterColumn>`,
          );
        }
      }
      // Simple filters (CT_Filters)
      for (const fi of af.filters ?? []) {
        const fcAttrs: Record<string, string | number | boolean | undefined> = {
          colId: fi.colId,
        };
        const filtersAttrs: Record<string, string | number | boolean | undefined> = {};
        if (fi.blank) filtersAttrs.blank = 1;
        if (fi.calendarType) filtersAttrs.calendarType = fi.calendarType;
        const valParts = (fi.values ?? []).map((v) => `<filter val="${escapeXml(v)}"/>`);
        inner.push(
          `<filterColumn${attrs(fcAttrs)}><filters${attrs(filtersAttrs)}>${valParts.join("")}</filters></filterColumn>`,
        );
      }
      if (af.sort && af.sort.length > 0) {
        const sortParts: string[] = [];
        for (const sc of af.sort) {
          const scAttrs: Record<string, string | number | boolean | undefined> = { ref: sc.ref };
          if (sc.descending) scAttrs.descending = 1;
          if (sc.sortBy) scAttrs.sortBy = sc.sortBy;
          if (sc.customList) scAttrs.customList = sc.customList;
          if (sc.iconId !== undefined) scAttrs.iconId = sc.iconId;
          sortParts.push(selfCloseElement("sortCondition", attrs(scAttrs)));
        }
        const ssAttrs: Record<string, string | number | boolean | undefined> = { ref: af.ref };
        if (af.sortState?.columnSort) ssAttrs.columnSort = 1;
        if (af.sortState?.caseSensitive) ssAttrs.caseSensitive = 1;
        if (af.sortState?.sortMethod) ssAttrs.sortMethod = af.sortState.sortMethod;
        inner.push(`<sortState${attrs(ssAttrs)}>${sortParts.join("")}</sortState>`);
      }
      // Color filters
      for (const cf of af.colorFilters ?? []) {
        const cfAttrs: Record<string, string | number | boolean | undefined> = {};
        if (cf.dxfId !== undefined) cfAttrs.dxfId = cf.dxfId;
        if (cf.cellColor === false) cfAttrs.cellColor = 0;
        inner.push(
          `<filterColumn colId="${cf.colId}"><colorFilter${attrs(cfAttrs)}/></filterColumn>`,
        );
      }
      // Icon filters
      for (const if_ of af.iconFilters ?? []) {
        const ifAttrs: Record<string, string | number | boolean | undefined> = {
          iconSet: if_.iconSet,
        };
        if (if_.iconId !== undefined) ifAttrs.iconId = if_.iconId;
        inner.push(
          `<filterColumn colId="${if_.colId}"><iconFilter${attrs(ifAttrs)}/></filterColumn>`,
        );
      }
      // Dynamic filters
      for (const df of af.dynamicFilters ?? []) {
        const dfAttrs: Record<string, string | number | boolean | undefined> = { type: df.type };
        if (df.val !== undefined) dfAttrs.val = df.val;
        if (df.maxVal !== undefined) dfAttrs.maxVal = df.maxVal;
        if (df.valIso !== undefined) dfAttrs.valIso = df.valIso;
        if (df.maxValIso !== undefined) dfAttrs.maxValIso = df.maxValIso;
        inner.push(
          `<filterColumn colId="${df.colId}"><dynamicFilter${attrs(dfAttrs)}/></filterColumn>`,
        );
      }
      // Date group filters
      for (const dg of af.dateGroupItems ?? []) {
        const dgAttrs: Record<string, string | number | boolean | undefined> = {
          dateTimeGrouping: dg.dateTimeGrouping,
        };
        if (dg.year !== undefined) dgAttrs.year = dg.year;
        if (dg.month !== undefined) dgAttrs.month = dg.month;
        if (dg.day !== undefined) dgAttrs.day = dg.day;
        if (dg.hour !== undefined) dgAttrs.hour = dg.hour;
        if (dg.minute !== undefined) dgAttrs.minute = dg.minute;
        if (dg.second !== undefined) dgAttrs.second = dg.second;
        inner.push(
          `<filterColumn colId="${dg.colId}"><dateGroupItem${attrs(dgAttrs)}/></filterColumn>`,
        );
      }
      if (inner.length > 0) {
        p.push(`<autoFilter ref="${af.ref}">`, ...inner, "</autoFilter>");
      } else {
        p.push(selfCloseElement("autoFilter", attrs({ ref: af.ref })));
      }
    }
  }

  // Merge cells
  if (mergeCells.length > 0) {
    p.push(`<mergeCells count="${mergeCells.length}">`);
    for (const mc of mergeCells) {
      const fromRef = defaultCellRef(mc.from.row, mc.from.col);
      const toRef = defaultCellRef(mc.to.row, mc.to.col);
      p.push(selfCloseElement("mergeCell", attrs({ ref: `${fromRef}:${toRef}` })));
    }
    p.push("</mergeCells>");
  }

  // Phonetic properties (after mergeCells per XSD sequence)
  if (opts.phoneticPr) {
    const pp = opts.phoneticPr;
    const ppAttrs: Record<string, string | number> = { fontId: pp.fontId };
    if (pp.type && pp.type !== "fullwidthKatakana") ppAttrs.type = pp.type;
    if (pp.alignment && pp.alignment !== "left") ppAttrs.alignment = pp.alignment;
    p.push(selfCloseElement("phoneticPr", attrs(ppAttrs)));
  }

  // Conditional formatting
  const conditionalFormats = opts.conditionalFormats ?? [];
  if (conditionalFormats.length > 0) {
    for (const cf of conditionalFormats) {
      p.push(`<conditionalFormatting sqref="${cf.sqref}">`);
      for (let ri = 0; ri < cf.rules.length; ri++) {
        const rule = cf.rules[ri];
        const ruleAttrs: Record<string, string | number | boolean | undefined> = {
          type: rule.type,
          priority: rule.priority ?? ri + 1,
        };
        if (rule.operator) ruleAttrs.operator = rule.operator;
        if (rule.dxfId !== undefined) ruleAttrs.dxfId = rule.dxfId;
        if (rule.stopIfTrue) ruleAttrs.stopIfTrue = 1;
        if (rule.timePeriod) ruleAttrs.timePeriod = rule.timePeriod;
        if (rule.rank !== undefined) ruleAttrs.rank = rule.rank;
        if (rule.equalAverage) ruleAttrs.equalAverage = 1;

        // Color scale
        if (rule.type === "colorScale" && rule.colorScale) {
          const cs = rule.colorScale;
          const inner: string[] = [];
          for (const v of cs.cfvo) {
            inner.push(buildCfvoXml(v));
          }
          for (const c of cs.colors) {
            inner.push(`<color rgb="FF${c}"/>`);
          }
          p.push(`<cfRule${attrs(ruleAttrs)}><colorScale>${inner.join("")}</colorScale></cfRule>`);
        }
        // Data bar
        else if (rule.type === "dataBar" && rule.dataBar) {
          const db = rule.dataBar;
          const inner: string[] = [];
          for (const v of db.cfvo) {
            inner.push(buildCfvoXml(v));
          }
          inner.push(`<color rgb="FF${db.color}"/>`);
          const dbAttrs: Record<string, string | number | boolean | undefined> = {};
          if (db.minLength !== undefined && db.minLength !== 10) dbAttrs.minLength = db.minLength;
          if (db.maxLength !== undefined && db.maxLength !== 90) dbAttrs.maxLength = db.maxLength;
          if (db.showValue === false) dbAttrs.showValue = 0;
          const attrStr = Object.keys(dbAttrs).length > 0 ? attrs(dbAttrs) : "";
          p.push(
            `<cfRule${attrs(ruleAttrs)}><dataBar${attrStr}>${inner.join("")}</dataBar></cfRule>`,
          );
        }
        // Icon set
        else if (rule.type === "iconSet" && rule.iconSet) {
          const is = rule.iconSet;
          const inner: string[] = [];
          for (const v of is.cfvo) {
            inner.push(buildCfvoXml(v));
          }
          const isAttrs: Record<string, string | number | boolean | undefined> = {};
          if (is.iconSet !== undefined && is.iconSet !== "3TrafficLights1")
            isAttrs.iconSet = is.iconSet;
          if (is.showValue === false) isAttrs.showValue = 0;
          if (is.percent === false) isAttrs.percent = 0;
          if (is.reverse) isAttrs.reverse = 1;
          const attrStr = Object.keys(isAttrs).length > 0 ? attrs(isAttrs) : "";
          p.push(
            `<cfRule${attrs(ruleAttrs)}><iconSet${attrStr}>${inner.join("")}</iconSet></cfRule>`,
          );
        }
        // Standard rules (cellIs, containsText, expression, top10, aboveAverage)
        else {
          if (rule.formulas && rule.formulas.length > 0) {
            const formulaParts = rule.formulas.map((f) => `<formula>${escapeXml(f)}</formula>`);
            p.push(`<cfRule${attrs(ruleAttrs)}>`, ...formulaParts, "</cfRule>");
          } else {
            p.push(selfCloseElement("cfRule", attrs(ruleAttrs)));
          }
        }
      }
      p.push("</conditionalFormatting>");
    }
  }

  // Data validations
  const dataValidations = opts.dataValidations ?? [];
  if (dataValidations.length > 0) {
    const dvContainerAttrs: Record<string, string | number | boolean | undefined> = {
      count: dataValidations.length,
    };
    if (opts.dataValidationsDisablePrompts) dvContainerAttrs.disablePrompts = 1;
    p.push(`<dataValidations${attrs(dvContainerAttrs)}>`);
    for (const dv of dataValidations) {
      const dvAttrs: Record<string, string | number | boolean | undefined> = { sqref: dv.sqref };
      if (dv.type && dv.type !== "none") dvAttrs.type = dv.type;
      if (dv.operator) dvAttrs.operator = dv.operator;
      if (dv.allowBlank) dvAttrs.allowBlank = 1;
      if (dv.showErrorMessage) dvAttrs.showErrorMessage = 1;
      if (dv.showInputMessage) dvAttrs.showInputMessage = 1;
      if (dv.errorTitle) dvAttrs.errorTitle = dv.errorTitle;
      if (dv.error) dvAttrs.error = dv.error;
      if (dv.promptTitle) dvAttrs.promptTitle = dv.promptTitle;
      if (dv.prompt) dvAttrs.prompt = dv.prompt;
      if (dv.errorStyle) dvAttrs.errorStyle = dv.errorStyle;
      if (dv.imeMode) dvAttrs.imeMode = dv.imeMode;
      if (dv.showDropDown) dvAttrs.showDropDown = 1;
      const inner: string[] = [];
      if (dv.formula1 !== undefined) inner.push(`<formula1>${escapeXml(dv.formula1)}</formula1>`);
      if (dv.formula2 !== undefined) inner.push(`<formula2>${escapeXml(dv.formula2)}</formula2>`);
      if (inner.length > 0) {
        p.push(`<dataValidation${attrs(dvAttrs)}>`, ...inner, "</dataValidation>");
      } else {
        p.push(selfCloseElement("dataValidation", attrs(dvAttrs)));
      }
    }
    p.push("</dataValidations>");
  }

  // Hyperlinks — r:id numbering must match worksheet rels order (compiler handles rels)
  const hyperlinks = opts.hyperlinks ?? [];
  if (hyperlinks.length > 0) {
    p.push("<hyperlinks>");
    let hlIdx = 0;
    for (const hl of hyperlinks) {
      const hlAttrs: Record<string, string | number | boolean | undefined> = { ref: hl.cell };
      if (hl.target.type === "external") {
        hlIdx++;
        hlAttrs["r:id"] = `rId${hlIdx}`;
      } else {
        hlAttrs.location = hl.target.location;
      }
      if (hl.tooltip) hlAttrs.tooltip = hl.tooltip;
      if (hl.display) hlAttrs.display = hl.display;
      p.push(selfCloseElement("hyperlink", attrs(hlAttrs)));
    }
    p.push("</hyperlinks>");
  }

  // Print options
  if (opts.printOptions) {
    const po = opts.printOptions;
    const poAttrs: Record<string, string | number | boolean | undefined> = {};
    if (po.horizontalCentered) poAttrs.horizontalCentered = 1;
    if (po.verticalCentered) poAttrs.verticalCentered = 1;
    if (po.headings) poAttrs.headings = 1;
    if (po.gridLines) poAttrs.gridLines = 1;
    if (po.gridLinesSet === false) poAttrs.gridLinesSet = 0;
    p.push(selfCloseElement("printOptions", attrs(poAttrs)));
  }

  p.push('<pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>');

  // Page setup
  if (opts.pageSetup) {
    const ps = opts.pageSetup;
    const psAttrs: Record<string, string | number | boolean | undefined> = {};
    if (ps.paperSize !== undefined) psAttrs.paperSize = ps.paperSize;
    if (ps.orientation && ps.orientation !== "default") psAttrs.orientation = ps.orientation;
    if (ps.scale !== undefined) psAttrs.scale = ps.scale;
    if (ps.fitToWidth !== undefined) psAttrs.fitToWidth = ps.fitToWidth;
    if (ps.fitToHeight !== undefined) psAttrs.fitToHeight = ps.fitToHeight;
    if (ps.pageOrder && ps.pageOrder !== "downThenOver") psAttrs.pageOrder = ps.pageOrder;
    if (ps.useFirstPageNumber) psAttrs.useFirstPageNumber = 1;
    if (ps.firstPageNumber !== undefined) psAttrs.firstPageNumber = ps.firstPageNumber;
    if (ps.paperHeight !== undefined) psAttrs.paperHeight = ps.paperHeight;
    if (ps.paperWidth !== undefined) psAttrs.paperWidth = ps.paperWidth;
    if (ps.usePrinterDefaults) psAttrs.usePrinterDefaults = 1;
    if (ps.blackAndWhite) psAttrs.blackAndWhite = 1;
    if (ps.draft) psAttrs.draft = 1;
    if (ps.cellComments && ps.cellComments !== "none") psAttrs.cellComments = ps.cellComments;
    if (ps.errors && ps.errors !== "displayed") psAttrs.errors = ps.errors;
    p.push(selfCloseElement("pageSetup", attrs(psAttrs)));
  }

  // Header/footer
  if (opts.headerFooter) {
    const hf = opts.headerFooter;
    const hfAttrs: Record<string, string | number | boolean | undefined> = {};
    if (hf.differentOddEven) hfAttrs.differentOddEven = 1;
    if (hf.differentFirst) hfAttrs.differentFirst = 1;
    if (hf.scaleWithDoc === false) hfAttrs.scaleWithDoc = 0;
    if (hf.alignWithMargins === false) hfAttrs.alignWithMargins = 0;
    const inner: string[] = [];
    if (hf.oddHeader) inner.push(`<oddHeader>${escapeXml(hf.oddHeader)}</oddHeader>`);
    if (hf.oddFooter) inner.push(`<oddFooter>${escapeXml(hf.oddFooter)}</oddFooter>`);
    if (hf.evenHeader) inner.push(`<evenHeader>${escapeXml(hf.evenHeader)}</evenHeader>`);
    if (hf.evenFooter) inner.push(`<evenFooter>${escapeXml(hf.evenFooter)}</evenFooter>`);
    if (hf.firstHeader) inner.push(`<firstHeader>${escapeXml(hf.firstHeader)}</firstHeader>`);
    if (hf.firstFooter) inner.push(`<firstFooter>${escapeXml(hf.firstFooter)}</firstFooter>`);
    if (inner.length > 0) {
      p.push(`<headerFooter${attrs(hfAttrs)}>`, ...inner, "</headerFooter>");
    } else if (hfAttrs.differentOddEven || hfAttrs.differentFirst) {
      p.push(selfCloseElement("headerFooter", attrs(hfAttrs)));
    }
  }

  // Drawing in header/footer (after headerFooter per XSD sequence)
  if (opts.drawingHF) {
    const dhf = opts.drawingHF;
    const dhfAttrs: Record<string, string | number | boolean | undefined> = { "r:id": dhf.rId };
    if (dhf.lho !== undefined) dhfAttrs.lho = dhf.lho;
    if (dhf.lhe !== undefined) dhfAttrs.lhe = dhf.lhe;
    if (dhf.lhf !== undefined) dhfAttrs.lhf = dhf.lhf;
    if (dhf.cho !== undefined) dhfAttrs.cho = dhf.cho;
    if (dhf.che !== undefined) dhfAttrs.che = dhf.che;
    if (dhf.chf !== undefined) dhfAttrs.chf = dhf.chf;
    if (dhf.rho !== undefined) dhfAttrs.rho = dhf.rho;
    if (dhf.rhe !== undefined) dhfAttrs.rhe = dhf.rhe;
    if (dhf.rhf !== undefined) dhfAttrs.rhf = dhf.rhf;
    if (dhf.lfo !== undefined) dhfAttrs.lfo = dhf.lfo;
    if (dhf.lfe !== undefined) dhfAttrs.lfe = dhf.lfe;
    if (dhf.lff !== undefined) dhfAttrs.lff = dhf.lff;
    if (dhf.cfo !== undefined) dhfAttrs.cfo = dhf.cfo;
    if (dhf.cfe !== undefined) dhfAttrs.cfe = dhf.cfe;
    if (dhf.cff !== undefined) dhfAttrs.cff = dhf.cff;
    if (dhf.rfo !== undefined) dhfAttrs.rfo = dhf.rfo;
    if (dhf.rfe !== undefined) dhfAttrs.rfe = dhf.rfe;
    if (dhf.rff !== undefined) dhfAttrs.rff = dhf.rff;
    p.push(selfCloseElement("drawingHF", attrs(dhfAttrs)));
  }

  // Legacy drawing in header/footer
  if (opts.legacyDrawingHF) {
    p.push(`<legacyDrawingHF r:id="${escapeXml(opts.legacyDrawingHF)}"/>`);
  }

  // Ignored errors (after headerFooter per XSD sequence)
  if (ignoredErrors.length > 0) {
    const ieParts: string[] = ["<ignoredErrors>"];
    for (const ie of ignoredErrors) {
      const ieAttrs: Record<string, string | number | boolean | undefined> = {
        sqref: ie.sqref,
      };
      if (ie.evalError) ieAttrs.evalError = 1;
      if (ie.twoDigitTextYear) ieAttrs.twoDigitTextYear = 1;
      if (ie.numberStoredAsText) ieAttrs.numberStoredAsText = 1;
      if (ie.formula) ieAttrs.formula = 1;
      if (ie.formulaRange) ieAttrs.formulaRange = 1;
      if (ie.unlockedFormula) ieAttrs.unlockedFormula = 1;
      if (ie.emptyCellReference) ieAttrs.emptyCellReference = 1;
      if (ie.listDataValidation) ieAttrs.listDataValidation = 1;
      if (ie.calculatedColumn) ieAttrs.calculatedColumn = 1;
      ieParts.push(selfCloseElement("ignoredError", attrs(ieAttrs)));
    }
    ieParts.push("</ignoredErrors>");
    p.push(ieParts.join(""));
  }

  // Background picture placeholder — compiler replaces with <picture r:id="rIdN"/>
  if (opts.backgroundImage) {
    p.push("<!--BACKGROUND_PICTURE-->");
  }

  // OLE objects (CT_OleObjects, after picture per XSD sequence)
  if (oleObjects.length > 0) {
    const oleParts: string[] = ["<oleObjects>"];
    for (const ole of oleObjects) {
      const oleAttrs: string[] = [`shapeId="${ole.shapeId}"`];
      if (ole.progId) oleAttrs.push(`progId="${escapeXml(ole.progId)}"`);
      if (ole.dvAspect && ole.dvAspect !== "DVASPECT_CONTENT")
        oleAttrs.push(`dvAspect="${ole.dvAspect}"`);
      if (ole.link) oleAttrs.push(`link="${escapeXml(ole.link)}"`);
      if (ole.oleUpdate) oleAttrs.push(`oleUpdate="${ole.oleUpdate}"`);
      if (ole.autoLoad) oleAttrs.push('autoLoad="1"');
      if (ole.rId) oleAttrs.push(`r:id="${escapeXml(ole.rId)}"`);
      // objectPr (CT_ObjectPr, optional child)
      if (ole.objectPr) {
        const opr = ole.objectPr;
        const oprAttrs: string[] = [];
        if (opr.locked === false) oprAttrs.push('locked="0"');
        if (opr.defaultSize === false) oprAttrs.push('defaultSize="0"');
        if (opr.print === false) oprAttrs.push('print="0"');
        if (opr.disabled) oprAttrs.push('disabled="1"');
        if (opr.uiObject) oprAttrs.push('uiObject="1"');
        if (opr.autoFill === false) oprAttrs.push('autoFill="0"');
        if (opr.autoLine === false) oprAttrs.push('autoLine="0"');
        if (opr.autoPict === false) oprAttrs.push('autoPict="0"');
        if (opr.macro) oprAttrs.push(`macro="${escapeXml(opr.macro)}"`);
        if (opr.altText) oprAttrs.push(`altText="${escapeXml(opr.altText)}"`);
        if (opr.dde) oprAttrs.push('dde="1"');
        if (opr.rId) oprAttrs.push(`r:id="${escapeXml(opr.rId)}"`);
        oleParts.push(
          `<oleObject ${oleAttrs.join(" ")}><objectPr${oprAttrs.length ? " " + oprAttrs.join(" ") : ""}/></oleObject>`,
        );
      } else {
        oleParts.push(`<oleObject ${oleAttrs.join(" ")}/>`);
      }
    }
    oleParts.push("</oleObjects>");
    p.push(oleParts.join(""));
  }

  // Controls (CT_Controls, after oleObjects per XSD sequence)
  if (controls.length > 0) {
    const ctrlParts: string[] = ["<controls>"];
    for (const c of controls) {
      const cAttrs: string[] = [`shapeId="${c.shapeId}"`, `r:id="${escapeXml(c.rId)}"`];
      if (c.name) cAttrs.push(`name="${escapeXml(c.name)}"`);
      // controlPr (optional)
      const prAttrs: string[] = [];
      if (c.locked === false) prAttrs.push('locked="0"');
      if (c.uiObject) prAttrs.push('uiObject="1"');
      if (c.recalcAlways) prAttrs.push('recalcAlways="1"');
      if (c.linkedCell) prAttrs.push(`linkedCell="${escapeXml(c.linkedCell)}"`);
      if (c.listFillRange) prAttrs.push(`listFillRange="${escapeXml(c.listFillRange)}"`);
      if (c.cf) prAttrs.push(`cf="${escapeXml(c.cf)}"`);
      if (prAttrs.length > 0) {
        ctrlParts.push(
          `<control ${cAttrs.join(" ")}><controlPr${prAttrs.length ? " " + prAttrs.join(" ") : ""}/></control>`,
        );
      } else {
        ctrlParts.push(`<control ${cAttrs.join(" ")}/>`);
      }
    }
    ctrlParts.push("</controls>");
    p.push(ctrlParts.join(""));
  }

  // Web publish items (CT_WebPublishItems, after controls per XSD sequence)
  if (webPublishItems.length > 0) {
    const wpParts: string[] = [`<webPublishItems count="${webPublishItems.length}">`];
    for (const wpi of webPublishItems) {
      const wpiAttrs: string[] = [
        `id="${wpi.id}"`,
        `divId="${escapeXml(wpi.divId)}"`,
        `sourceType="${wpi.sourceType}"`,
        `destinationFile="${escapeXml(wpi.destinationFile)}"`,
      ];
      if (wpi.sourceRef) wpiAttrs.push(`sourceRef="${escapeXml(wpi.sourceRef)}"`);
      if (wpi.sourceObject) wpiAttrs.push(`sourceObject="${escapeXml(wpi.sourceObject)}"`);
      if (wpi.title) wpiAttrs.push(`title="${escapeXml(wpi.title)}"`);
      if (wpi.autoRepublish) wpiAttrs.push('autoRepublish="1"');
      wpParts.push(`<webPublishItem ${wpiAttrs.join(" ")}/>`);
    }
    wpParts.push("</webPublishItems>");
    p.push(wpParts.join(""));
  }

  // Extension list (extLst, last per XSD sequence)
  if (opts.ext) {
    p.push(`<extLst>${opts.ext}</extLst>`);
  }

  p.push("</worksheet>");
  return p.join("");
}

// ── Stringify helpers ──

function buildCfvoXml(cfvo: CfvoOptions): string {
  const a: Record<string, string | number | boolean | undefined> = { type: cfvo.type };
  if (cfvo.val !== undefined) a.val = cfvo.val;
  if (cfvo.gte === false) a.gte = 0;
  return `<cfvo${attrs(a)}/>`;
}

function buildSheetViewAttrs(sv?: SheetViewOptions): string {
  const svMap: Record<string, string | number | boolean | undefined> = {
    workbookViewId: 0,
  };
  if (sv?.tabSelected !== undefined) svMap.tabSelected = sv.tabSelected ? 1 : 0;
  else svMap.tabSelected = 1;
  if (sv?.showGridLines === false) svMap.showGridLines = 0;
  if (sv?.showRowColHeaders === false) svMap.showRowColHeaders = 0;
  if (sv?.showZeros === false) svMap.showZeros = 0;
  if (sv?.zoomScale !== undefined) svMap.zoomScale = sv.zoomScale;
  if (sv?.rightToLeft) svMap.rightToLeft = 1;
  if (sv?.windowProtection) svMap.windowProtection = 1;
  if (sv?.showFormulas) svMap.showFormulas = 1;
  if (sv?.showRuler === false) svMap.showRuler = 0;
  if (sv?.showOutlineSymbols === false) svMap.showOutlineSymbols = 0;
  if (sv?.defaultGridColor === false) svMap.defaultGridColor = 0;
  if (sv?.showWhiteSpace === false) svMap.showWhiteSpace = 0;
  if (sv?.view) svMap.view = sv.view;
  if (sv?.colorId !== undefined) svMap.colorId = sv.colorId;
  if (sv?.zoomScaleNormal !== undefined) svMap.zoomScaleNormal = sv.zoomScaleNormal;
  if (sv?.zoomScaleSheetLayoutView !== undefined)
    svMap.zoomScaleSheetLayoutView = sv.zoomScaleSheetLayoutView;
  if (sv?.zoomScalePageLayoutView !== undefined)
    svMap.zoomScalePageLayoutView = sv.zoomScalePageLayoutView;
  return attrs(svMap);
}

function buildSelectionXml(sel: SelectionOptions): string {
  const selAttrs: Record<string, string | number | boolean | undefined> = {};
  if (sel.pane) selAttrs.pane = sel.pane;
  if (sel.activeCell) selAttrs.activeCell = sel.activeCell;
  if (sel.activeCellId !== undefined) selAttrs.activeCellId = sel.activeCellId;
  if (sel.sqref) selAttrs.sqref = sel.sqref;
  return `<selection${attrs(selAttrs)}/>`;
}

function buildPivotSelectionXml(_ps: PivotSelectionOptions): string {
  // pivotSelection is optional; omit if no meaningful pivotArea can be constructed.
  // An empty <pivotArea/> causes Excel to reject the file.
  return "";
}

function hashPassword(password: string): string {
  let hash = 0;
  for (let i = 0; i < password.length; i++) {
    const c = password.charCodeAt(i);
    hash = ((hash >> 14) & 1) + ((hash << 1) & 0x7fff);
    hash ^= c;
    hash = hash & 0x4000 ? hash ^ 0x1 : hash;
  }
  hash = ((hash >> 14) & 1) + ((hash << 1) & 0x7fff);
  hash = ((hash >> 14) & 1) + ((hash << 1) & 0x7fff);
  hash ^= password.length;
  return hash.toString(16).toUpperCase().padStart(4, "0");
}

function buildFormulaString(fOpts: FormulaOptions): string {
  const fAttrs: Record<string, string | number | boolean | undefined> = {};
  if (fOpts.type && fOpts.type !== FormulaType.NORMAL) fAttrs.t = fOpts.type;
  if (fOpts.reference) fAttrs.ref = fOpts.reference;
  if (fOpts.sharedIndex !== undefined) fAttrs.si = fOpts.sharedIndex;
  if (fOpts.aca) fAttrs.aca = 1;
  if (fOpts.dt2D) fAttrs.dt2D = 1;
  if (fOpts.dtr) fAttrs.dtr = 1;
  if (fOpts.del1) fAttrs.del1 = 1;
  if (fOpts.del2) fAttrs.del2 = 1;
  if (fOpts.r1) fAttrs.r1 = fOpts.r1;
  if (fOpts.r2) fAttrs.r2 = fOpts.r2;
  if (fOpts.ca) fAttrs.ca = 1;
  if (fOpts.bx) fAttrs.bx = 1;

  const hasContent = fOpts.formula !== undefined && fOpts.formula !== "";

  if (hasContent) {
    return `<f${attrs(fAttrs)}>${escapeXml(fOpts.formula)}</f>`;
  }
  if (Object.keys(fAttrs).length > 0) {
    return selfCloseElement("f", attrs(fAttrs));
  }
  return "";
}

function buildCellString(
  ref: string,
  cell: CellOptions,
  sharedStrings?: SharedStrings,
  styles?: Styles,
): string {
  const cellAttrs: Record<string, string | number | boolean | undefined> = { r: ref };

  // Resolve style
  if (cell.style !== undefined && styles) {
    cellAttrs.s = styles.register(cell.style);
  } else if (cell.styleIndex !== undefined) {
    cellAttrs.s = cell.styleIndex;
  }

  const value = cell.value;

  // Formula path — formula takes precedence; value is the cached result.
  if (cell.formula) {
    const fStr = buildFormulaString(cell.formula);
    let vStr = "";
    if (value === null || value === undefined) {
      return `<c${attrs(cellAttrs)}>${fStr}</c>`;
    }
    if (typeof value === "number") {
      vStr = `<v>${value}</v>`;
    } else if (typeof value === "boolean") {
      cellAttrs.t = "b";
      vStr = `<v>${value ? 1 : 0}</v>`;
    } else if (typeof value === "string") {
      cellAttrs.t = "str";
      vStr = `<v>${escapeXml(value)}</v>`;
    } else if (value instanceof Date) {
      vStr = `<v>${dateToSerialNumber(value)}</v>`;
    }
    if (vStr) {
      return `<c${attrs(cellAttrs)}>${fStr}${vStr}</c>`;
    }
    return `<c${attrs(cellAttrs)}>${fStr}</c>`;
  }

  if (value === null || value === undefined) {
    if (cell.styleIndex !== undefined) {
      return selfCloseElement("c", attrs(cellAttrs));
    }
    return "";
  }

  // Rich text value (RichTextOptions)
  if (typeof value === "object" && !(value instanceof Date)) {
    if (sharedStrings) {
      cellAttrs.t = "s";
      const idx = sharedStrings.registerRich(value);
      return `<c${attrs(cellAttrs)}><v>${idx}</v></c>`;
    }
    cellAttrs.t = "inlineStr";
    return `<c${attrs(cellAttrs)}><is>${buildRstXml(value)}</is></c>`;
  }

  if (typeof value === "string") {
    if (sharedStrings) {
      cellAttrs.t = "s";
      const idx = sharedStrings.register(value);
      return `<c${attrs(cellAttrs)}><v>${idx}</v></c>`;
    }
    cellAttrs.t = "inlineStr";
    return `<c${attrs(cellAttrs)}><is><t>${escapeXml(value)}</t></is></c>`;
  }

  if (typeof value === "number") {
    return `<c${attrs(cellAttrs)}><v>${value}</v></c>`;
  }

  if (typeof value === "boolean") {
    cellAttrs.t = "b";
    return `<c${attrs(cellAttrs)}><v>${value ? 1 : 0}</v></c>`;
  }

  if (value instanceof Date) {
    const serial = dateToSerialNumber(value);
    return `<c${attrs(cellAttrs)}><v>${serial}</v></c>`;
  }

  return "";
}

function defaultCellRef(row: number, col: number): string {
  return columnToLetter(col) + row;
}

function columnToLetter(col: number): string {
  let result = "";
  let n = col;
  while (n > 0) {
    const remainder = (n - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

function dateToSerialNumber(date: Date): number {
  const epoch = new Date(1899, 11, 30);
  const msPerDay = 86400000;
  return (date.getTime() - epoch.getTime()) / msPerDay;
}

// ── Parse helpers ──

function parseCfvo(el: XmlElement): Record<string, unknown> {
  const result: Record<string, unknown> = {};
  result.type = attr(el, "type") ?? "num";
  const val = attr(el, "val");
  if (val !== undefined) result.val = isNaN(Number(val)) ? val : Number(val);
  if (attr(el, "gte") === "0") result.gte = false;
  return result;
}

function parseCellRef(ref: string): { row: number; col: number } | undefined {
  const match = ref.match(/^([A-Z]+)(\d+)$/);
  if (!match) return undefined;
  const colStr = match[1];
  const row = parseInt(match[2], 10);
  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }
  return { row, col };
}
