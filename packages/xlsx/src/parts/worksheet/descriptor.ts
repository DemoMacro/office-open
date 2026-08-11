/**
 * Worksheet — descriptor for xl/worksheets/sheet{n}.xml.
 *
 * stringify() intentionally throws: the compiler calls stringifyWorksheet()
 * directly (it needs the SharedStrings/Styles accumulators). This descriptor
 * exists for the read path only.
 *
 * @module
 */
import type { PositiveUniversalMeasure } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrMeasure, attrNum, findChild, stringify, textOf } from "@office-open/xml";

import type { XlsxReadContext } from "../../context";
import { parseAutoFilter } from "../auto-filter";
import { parseCellRef, parseCfvo, parsePageBreaks } from "./parse";
import type {
  CellOptions,
  CellWatchOptions,
  CfvoOptions,
  ColumnOptions,
  ConditionalFormatOperator,
  ConditionalFormatOptions,
  ConditionalFormatRule,
  ConditionalFormatType,
  ControlOptions,
  CustomPropertyOptions,
  CustomSheetViewOptions,
  DataConsolidateOptions,
  DataValidationOperator,
  DataValidationOptions,
  DataValidationType,
  DrawingHfOptions,
  FormulaOptions,
  FormulaType,
  FreezePaneOptions,
  HeaderFooterOptions,
  HyperlinkOptions,
  HyperlinkTarget,
  IconSetOptions,
  IconSetType,
  IgnoredErrorOptions,
  MergeCellOptions,
  OleObjectOptions,
  OleObjectPropertiesOptions,
  PageMarginsOptions,
  PageOrientation,
  PageSetupOptions,
  PhoneticPropertiesOptions,
  PrintOptions,
  ProtectedRangeOptions,
  RowOptions,
  ScenarioCellOptions,
  ScenarioDefinition,
  ScenarioOptions,
  SheetCalculationPropertiesOptions,
  SheetFormatPropertiesOptions,
  SheetPropertiesOptions,
  SheetProtectionOptions,
  SheetViewOptions,
  TabColorOptions,
  WebPublishItemOptions,
  WorksheetOptions,
} from "./types";

const DRAWING_HF_OFFSET_KEYS = [
  "lho",
  "lhe",
  "lhf",
  "cho",
  "che",
  "chf",
  "rho",
  "rhe",
  "rhf",
  "lfo",
  "lfe",
  "lff",
  "cfo",
  "cfe",
  "cff",
  "rfo",
  "rfe",
  "rff",
] as const satisfies readonly (keyof DrawingHfOptions)[];

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

  parse(el, ctx) {
    const result: Partial<WorksheetOptions> = {};
    let pageSetUpPrCache: Partial<PageSetupOptions> | undefined;

    // Resolve shared strings from context (XlsxReadContext)
    const strings: string[] =
      ctx && "sharedStrings" in ctx ? (ctx as XlsxReadContext).sharedStrings : [];

    // Sheet properties
    const sheetPrEl = findChild(el, "sheetPr");
    if (sheetPrEl) {
      const sp: SheetPropertiesOptions = {};
      if (String(attr(sheetPrEl, "syncHorizontal")) === "1") sp.syncHorizontal = true;
      if (String(attr(sheetPrEl, "syncVertical")) === "1") sp.syncVertical = true;
      if (attr(sheetPrEl, "syncRef")) sp.syncRef = attr(sheetPrEl, "syncRef");
      if (String(attr(sheetPrEl, "transitionEvaluation")) === "1") sp.transitionEvaluation = true;
      if (String(attr(sheetPrEl, "transitionEntry")) === "1") sp.transitionEntry = true;
      if (String(attr(sheetPrEl, "published")) === "1") sp.published = true;
      if (String(attr(sheetPrEl, "filterMode")) === "1") sp.filterMode = true;
      if (String(attr(sheetPrEl, "enableFormatConditionsCalculation")) === "1")
        sp.enableFormatConditionsCalculation = true;

      const outlinePr = findChild(sheetPrEl, "outlinePr");
      if (outlinePr) {
        if (String(attr(outlinePr, "applyStyles")) === "1") sp.outlineApplyStyles = true;
        if (String(attr(outlinePr, "showOutlineSymbols")) === "0") sp.outlineShowSymbols = false;
        if (String(attr(outlinePr, "summaryBelow")) === "0") sp.outlineSummaryBelow = false;
        if (String(attr(outlinePr, "summaryRight")) === "0") sp.outlineSummaryRight = false;
      }

      // pageSetUpPr (inside sheetPr) — stash on result.pageSetup; merged into
      // the <pageSetup> parse below, which owns result.pageSetup.
      const pageSetUpPr = findChild(sheetPrEl, "pageSetUpPr");
      if (pageSetUpPr) {
        const psup: Partial<PageSetupOptions> = {};
        if (String(attr(pageSetUpPr, "fitToPage")) === "1") psup.fitToPage = true;
        if (String(attr(pageSetUpPr, "autoPageBreaks")) === "1") psup.autoPageBreaks = true;
        if (Object.keys(psup).length > 0) pageSetUpPrCache = psup;
      }
      if (Object.keys(sp).length > 0) result.sheetPr = sp;

      // Tab color
      const tabColorEl = findChild(sheetPrEl, "tabColor");
      if (tabColorEl) {
        const tc: TabColorOptions = {};
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
        const sv: SheetViewOptions = {};
        if (String(attr(svEl, "showGridLines")) === "0") sv.showGridLines = false;
        if (String(attr(svEl, "showRowColHeaders")) === "0") sv.showRowColHeaders = false;
        if (String(attr(svEl, "showZeros")) === "0") sv.showZeros = false;
        const zs = attrNum(svEl, "zoomScale");
        if (zs !== undefined) sv.zoomScale = zs;
        if (attr(svEl, "tabSelected") !== undefined)
          sv.tabSelected = attr(svEl, "tabSelected") !== "0";
        if (String(attr(svEl, "rightToLeft")) === "1") sv.rightToLeft = true;
        if (String(attr(svEl, "windowProtection")) === "1") sv.windowProtection = true;
        if (String(attr(svEl, "showFormulas")) === "1") sv.showFormulas = true;
        if (String(attr(svEl, "showRuler")) === "0") sv.showRuler = false;
        if (String(attr(svEl, "showOutlineSymbols")) === "0") sv.showOutlineSymbols = false;
        if (String(attr(svEl, "defaultGridColor")) === "0") sv.defaultGridColor = false;
        if (String(attr(svEl, "showWhiteSpace")) === "0") sv.showWhiteSpace = false;
        const viewVal = attr(svEl, "view");
        if (viewVal) sv.view = viewVal as SheetViewOptions["view"];
        const colorId = attrNum(svEl, "colorId");
        if (colorId !== undefined) sv.colorId = colorId;
        const zsn = attrNum(svEl, "zoomScaleNormal");
        if (zsn !== undefined) sv.zoomScaleNormal = zsn;
        const zssl = attrNum(svEl, "zoomScaleSheetLayoutView");
        if (zssl !== undefined) sv.zoomScaleSheetLayoutView = zssl;
        const zspl = attrNum(svEl, "zoomScalePageLayoutView");
        if (zspl !== undefined) sv.zoomScalePageLayoutView = zspl;
        result.sheetView = sv;

        // Freeze pane
        const paneEl = findChild(svEl, "pane");
        if (paneEl && attr(paneEl, "state") === "frozen") {
          const fp: FreezePaneOptions = {};
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
      const sfp: SheetFormatPropertiesOptions = {};
      const bcw = attrNum(sfpEl, "baseColWidth");
      if (bcw !== undefined) sfp.baseColWidth = bcw;
      const dcw = attrNum(sfpEl, "defaultColWidth");
      if (dcw !== undefined) sfp.defaultColWidth = dcw;
      const drh = attrNum(sfpEl, "defaultRowHeight");
      if (drh !== undefined) sfp.defaultRowHeight = drh;
      if (String(attr(sfpEl, "zeroHeight")) === "1") sfp.zeroHeight = true;
      if (String(attr(sfpEl, "thickTop")) === "1") sfp.thickTop = true;
      if (String(attr(sfpEl, "thickBottom")) === "1") sfp.thickBottom = true;
      const olr = attrNum(sfpEl, "outlineLevelRow");
      if (olr !== undefined) sfp.outlineLevelRow = olr;
      const olc = attrNum(sfpEl, "outlineLevelCol");
      if (olc !== undefined) sfp.outlineLevelCol = olc;
      result.sheetFormatPr = sfp;
    }

    // Dimension
    const dimensionEl = findChild(el, "dimension");
    if (dimensionEl) {
      const ref = attr(dimensionEl, "ref");
      if (ref) result.dimension = ref;
    }

    // Page margins
    const pageMarginsEl = findChild(el, "pageMargins");
    if (pageMarginsEl) {
      const pm: PageMarginsOptions = {};
      const pmL = attrNum(pageMarginsEl, "left");
      if (pmL !== undefined) pm.left = pmL;
      const pmR = attrNum(pageMarginsEl, "right");
      if (pmR !== undefined) pm.right = pmR;
      const pmT = attrNum(pageMarginsEl, "top");
      if (pmT !== undefined) pm.top = pmT;
      const pmB = attrNum(pageMarginsEl, "bottom");
      if (pmB !== undefined) pm.bottom = pmB;
      const pmH = attrNum(pageMarginsEl, "header");
      if (pmH !== undefined) pm.header = pmH;
      const pmF = attrNum(pageMarginsEl, "footer");
      if (pmF !== undefined) pm.footer = pmF;
      result.pageMargins = pm;
    }

    // Columns
    const colsEl = findChild(el, "cols");
    if (colsEl) {
      const columns: ColumnOptions[] = [];
      for (const colEl of colsEl.elements ?? []) {
        if (colEl.name !== "col") continue;
        const col: ColumnOptions = {
          min: attrNum(colEl, "min") ?? 0,
          max: attrNum(colEl, "max") ?? 0,
        };
        const w = attrNum(colEl, "width");
        if (w !== undefined) col.width = w;
        if (String(attr(colEl, "hidden")) === "1") col.hidden = true;
        if (String(attr(colEl, "customWidth")) === "1") col.customWidth = true;
        const ol = attrNum(colEl, "outlineLevel");
        if (ol !== undefined) col.outlineLevel = ol;
        if (String(attr(colEl, "collapsed")) === "1") col.collapsed = true;
        if (String(attr(colEl, "bestFit")) === "1") col.bestFit = true;
        if (String(attr(colEl, "phonetic")) === "1") col.phonetic = true;
        columns.push(col);
      }
      if (columns.length > 0) result.columns = columns;
    }

    // Sheet protection
    const protEl = findChild(el, "sheetProtection");
    if (protEl?.attributes) {
      const prot: SheetProtectionOptions = {};
      if (attr(protEl, "password")) prot.password = attr(protEl, "password");
      if (attr(protEl, "algorithmName")) prot.algorithmName = attr(protEl, "algorithmName");
      if (attr(protEl, "hashValue")) prot.hashValue = attr(protEl, "hashValue");
      if (attr(protEl, "saltValue")) prot.saltValue = attr(protEl, "saltValue");
      if (attrNum(protEl, "spinCount") !== undefined) prot.spinCount = attrNum(protEl, "spinCount");
      if (String(attr(protEl, "sheet")) === "1") prot.sheet = true;
      if (String(attr(protEl, "objects")) === "1") prot.objects = true;
      if (String(attr(protEl, "scenarios")) === "1") prot.scenarios = true;
      if (String(attr(protEl, "formatCells")) === "0") prot.formatCells = false;
      if (String(attr(protEl, "formatColumns")) === "0") prot.formatColumns = false;
      if (String(attr(protEl, "formatRows")) === "0") prot.formatRows = false;
      if (String(attr(protEl, "insertColumns")) === "0") prot.insertColumns = false;
      if (String(attr(protEl, "insertRows")) === "0") prot.insertRows = false;
      if (String(attr(protEl, "insertHyperlinks")) === "0") prot.insertHyperlinks = false;
      if (String(attr(protEl, "deleteColumns")) === "0") prot.deleteColumns = false;
      if (String(attr(protEl, "deleteRows")) === "0") prot.deleteRows = false;
      if (String(attr(protEl, "selectLockedCells")) === "1") prot.selectLockedCells = true;
      if (String(attr(protEl, "sort")) === "0") prot.sort = false;
      if (String(attr(protEl, "autoFilter")) === "0") prot.autoFilter = false;
      if (String(attr(protEl, "pivotTables")) === "0") prot.pivotTables = false;
      if (String(attr(protEl, "selectUnlockedCells")) === "1") prot.selectUnlockedCells = true;
      result.protection = prot;
    }

    // Protected ranges
    const prEl = findChild(el, "protectedRanges");
    if (prEl) {
      const ranges: ProtectedRangeOptions[] = [];
      for (const rEl of prEl.elements ?? []) {
        if (rEl.name !== "protectedRange") continue;
        const r: ProtectedRangeOptions = {
          sqref: attr(rEl, "sqref") ?? "",
          name: attr(rEl, "name") ?? "",
        };
        if (attr(rEl, "password")) r.password = attr(rEl, "password");
        if (attr(rEl, "algorithmName")) r.algorithmName = attr(rEl, "algorithmName");
        if (attr(rEl, "hashValue")) r.hashValue = attr(rEl, "hashValue");
        if (attr(rEl, "saltValue")) r.saltValue = attr(rEl, "saltValue");
        const spinCount = attrNum(rEl, "spinCount");
        if (spinCount !== undefined) r.spinCount = spinCount;
        const sdEl = findChild(rEl, "securityDescriptor");
        if (sdEl) r.securityDescriptor = textOf(sdEl) ?? undefined;
        ranges.push(r);
      }
      if (ranges.length > 0) result.protectedRanges = ranges;
    }

    // Auto filter (CT_AutoFilter is shared with table — logic in auto-filter.ts)
    const afEl = findChild(el, "autoFilter");
    if (afEl) {
      result.autoFilter = parseAutoFilter(afEl);
    }

    // Merge cells
    const mcEl = findChild(el, "mergeCells");
    if (mcEl) {
      const merges: MergeCellOptions[] = [];
      for (const mEl of mcEl.elements ?? []) {
        if (mEl.name !== "mergeCell") continue;
        const ref = attr(mEl, "ref") ?? "";
        const parts = ref.split(":");
        if (parts.length === 2) {
          const from = parseCellRef(parts[0] ?? "");
          const to = parseCellRef(parts[1] ?? "");
          if (from && to) merges.push({ from, to });
        }
      }
      if (merges.length > 0) result.mergeCells = merges;
    }

    // Conditional formatting
    const cfEls = el.elements?.filter((e) => e.name === "conditionalFormatting") ?? [];
    if (cfEls.length > 0) {
      const cfs: ConditionalFormatOptions[] = [];
      for (const cfEl of cfEls) {
        const sqref = attr(cfEl, "sqref") ?? "";
        const rules: ConditionalFormatRule[] = [];
        for (const ruleEl of cfEl.elements ?? []) {
          if (ruleEl.name !== "cfRule") continue;
          const rule: ConditionalFormatRule = {
            type: attr(ruleEl, "type") as ConditionalFormatType,
            priority: attrNum(ruleEl, "priority") ?? 1,
          };
          const opVal = attr(ruleEl, "operator");
          if (opVal) rule.operator = opVal as ConditionalFormatOperator;
          const dxfId = attrNum(ruleEl, "dxfId");
          if (dxfId !== undefined) rule.dxfId = dxfId;
          if (String(attr(ruleEl, "stopIfTrue")) === "1") rule.stopIfTrue = true;
          const tpVal = attr(ruleEl, "timePeriod");
          if (tpVal) rule.timePeriod = tpVal as ConditionalFormatRule["timePeriod"];
          const rank = attrNum(ruleEl, "rank");
          if (rank !== undefined) rule.rank = rank;
          if (String(attr(ruleEl, "equalAverage")) === "1") rule.equalAverage = true;

          // Color scale
          const csEl = findChild(ruleEl, "colorScale");
          if (csEl) {
            const cfvo: CfvoOptions[] = [];
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
            const cfvo: CfvoOptions[] = [];
            let color = "";
            for (const child of dbEl.elements ?? []) {
              if (child.name === "cfvo") cfvo.push(parseCfvo(child));
              if (child.name === "color") {
                const rgb = attr(child, "rgb");
                if (rgb) color = rgb.length === 8 ? rgb.slice(2) : rgb;
              }
            }
            rule.dataBar = { cfvo: cfvo as [CfvoOptions, CfvoOptions], color };
          }

          // Icon set
          const isEl = findChild(ruleEl, "iconSet");
          if (isEl) {
            const cfvo: CfvoOptions[] = [];
            for (const child of isEl.elements ?? []) {
              if (child.name === "cfvo") cfvo.push(parseCfvo(child));
            }
            const iconSet: IconSetOptions = { cfvo };
            const isVal = attr(isEl, "iconSet");
            if (isVal) iconSet.iconSet = isVal as IconSetType;
            if (String(attr(isEl, "showValue")) === "0") iconSet.showValue = false;
            if (String(attr(isEl, "percent")) === "0") iconSet.percent = false;
            if (String(attr(isEl, "reverse")) === "1") iconSet.reverse = true;
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
      const dvs: DataValidationOptions[] = [];
      for (const dEl of dvEl.elements ?? []) {
        if (dEl.name !== "dataValidation") continue;
        const dv: DataValidationOptions = { sqref: attr(dEl, "sqref") ?? "" };
        const typeVal = attr(dEl, "type");
        if (typeVal) dv.type = typeVal as DataValidationType;
        const opVal = attr(dEl, "operator");
        if (opVal) dv.operator = opVal as DataValidationOperator;
        if (String(attr(dEl, "allowBlank")) === "1") dv.allowBlank = true;
        if (String(attr(dEl, "showErrorMessage")) === "1") dv.showErrorMessage = true;
        if (String(attr(dEl, "showInputMessage")) === "1") dv.showInputMessage = true;
        if (attr(dEl, "errorTitle")) dv.errorTitle = attr(dEl, "errorTitle");
        if (attr(dEl, "error")) dv.error = attr(dEl, "error");
        if (attr(dEl, "promptTitle")) dv.promptTitle = attr(dEl, "promptTitle");
        if (attr(dEl, "prompt")) dv.prompt = attr(dEl, "prompt");
        const esVal = attr(dEl, "errorStyle");
        if (esVal) dv.errorStyle = esVal as DataValidationOptions["errorStyle"];
        const imVal = attr(dEl, "imeMode");
        if (imVal) dv.imeMode = imVal as DataValidationOptions["imeMode"];
        if (String(attr(dEl, "showDropDown")) === "1") dv.showDropDown = true;

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
      const hyperlinks: HyperlinkOptions[] = [];
      for (const hEl of hlEl.elements ?? []) {
        if (hEl.name !== "hyperlink") continue;
        const rId = hEl.attributes?.["r:id"] as string | undefined;
        const location = attr(hEl, "location");
        const target: HyperlinkTarget = rId
          ? { type: "external", url: rId }
          : { type: "internal", location: location ?? "" };
        const hl: HyperlinkOptions = { cell: attr(hEl, "ref") ?? "", target };
        if (attr(hEl, "tooltip")) hl.tooltip = attr(hEl, "tooltip");
        if (attr(hEl, "display")) hl.display = attr(hEl, "display");
        hyperlinks.push(hl);
      }
      result.hyperlinks = hyperlinks;
    }

    // Print options
    const poEl = findChild(el, "printOptions");
    if (poEl) {
      const po: PrintOptions = {};
      if (String(attr(poEl, "horizontalCentered")) === "1") po.horizontalCentered = true;
      if (String(attr(poEl, "verticalCentered")) === "1") po.verticalCentered = true;
      if (String(attr(poEl, "headings")) === "1") po.headings = true;
      if (String(attr(poEl, "gridLines")) === "1") po.gridLines = true;
      if (String(attr(poEl, "gridLinesSet")) === "0") po.gridLinesSet = false;
      result.printOptions = po;
    }

    // Page setup
    const psEl = findChild(el, "pageSetup");
    if (psEl) {
      const ps: PageSetupOptions = {};
      const pz = attrNum(psEl, "paperSize");
      if (pz !== undefined) ps.paperSize = pz;
      const ph = attrMeasure(psEl, "paperHeight");
      if (ph !== undefined) ps.paperHeight = ph as number | PositiveUniversalMeasure;
      const pw = attrMeasure(psEl, "paperWidth");
      if (pw !== undefined) ps.paperWidth = pw as number | PositiveUniversalMeasure;
      const orientVal = attr(psEl, "orientation");
      if (orientVal) ps.orientation = orientVal as PageOrientation;
      const sc = attrNum(psEl, "scale");
      if (sc !== undefined) ps.scale = sc;
      const ftw = attrNum(psEl, "fitToWidth");
      if (ftw !== undefined) ps.fitToWidth = ftw;
      const fth = attrNum(psEl, "fitToHeight");
      if (fth !== undefined) ps.fitToHeight = fth;
      const pageOrderVal = attr(psEl, "pageOrder");
      if (pageOrderVal) ps.pageOrder = pageOrderVal as PageSetupOptions["pageOrder"];
      if (String(attr(psEl, "useFirstPageNumber")) === "1") ps.useFirstPageNumber = true;
      const fpn = attrNum(psEl, "firstPageNumber");
      if (fpn !== undefined) ps.firstPageNumber = fpn;
      if (pageSetUpPrCache) Object.assign(ps, pageSetUpPrCache);
      result.pageSetup = ps;
    } else if (pageSetUpPrCache) {
      result.pageSetup = pageSetUpPrCache;
    }

    // Header/footer
    const hfEl = findChild(el, "headerFooter");
    if (hfEl) {
      const hf: HeaderFooterOptions = {};
      if (String(attr(hfEl, "differentOddEven")) === "1") hf.differentOddEven = true;
      if (String(attr(hfEl, "differentFirst")) === "1") hf.differentFirst = true;
      if (String(attr(hfEl, "scaleWithDoc")) === "0") hf.scaleWithDoc = false;
      if (String(attr(hfEl, "alignWithMargins")) === "0") hf.alignWithMargins = false;
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
      const errors: IgnoredErrorOptions[] = [];
      for (const eEl of ieEl.elements ?? []) {
        if (eEl.name !== "ignoredError") continue;
        const ie: IgnoredErrorOptions = { sqref: attr(eEl, "sqref") ?? "" };
        if (String(attr(eEl, "evalError")) === "1") ie.evalError = true;
        if (String(attr(eEl, "twoDigitTextYear")) === "1") ie.twoDigitTextYear = true;
        if (String(attr(eEl, "numberStoredAsText")) === "1") ie.numberStoredAsText = true;
        if (String(attr(eEl, "formula")) === "1") ie.formula = true;
        if (String(attr(eEl, "formulaRange")) === "1") ie.formulaRange = true;
        if (String(attr(eEl, "unlockedFormula")) === "1") ie.unlockedFormula = true;
        if (String(attr(eEl, "emptyCellReference")) === "1") ie.emptyCellReference = true;
        if (String(attr(eEl, "listDataValidation")) === "1") ie.listDataValidation = true;
        if (String(attr(eEl, "calculatedColumn")) === "1") ie.calculatedColumn = true;
        errors.push(ie);
      }
      result.ignoredErrors = errors;
    }

    // Phonetic properties
    const ppEl = findChild(el, "phoneticPr");
    if (ppEl) {
      const pp: PhoneticPropertiesOptions = { fontId: attrNum(ppEl, "fontId") ?? 0 };
      const ppType = attr(ppEl, "type");
      if (ppType) pp.type = ppType as PhoneticPropertiesOptions["type"];
      const ppAlign = attr(ppEl, "alignment");
      if (ppAlign) pp.alignment = ppAlign as PhoneticPropertiesOptions["alignment"];
      result.phoneticPr = pp;
    }

    // Sheet calc properties
    const scEl = findChild(el, "sheetCalcPr");
    if (scEl) {
      const sc: SheetCalculationPropertiesOptions = {};
      if (String(attr(scEl, "fullCalcOnLoad")) === "1") sc.fullCalcOnLoad = true;
      result.sheetCalcPr = sc;
    }

    // Sheet data (rows and cells)
    const sheetDataEl = findChild(el, "sheetData");
    if (sheetDataEl) {
      const rows: RowOptions[] = [];
      for (const rowEl of sheetDataEl.elements ?? []) {
        if (rowEl.name !== "row") continue;
        const row: RowOptions = {};
        const rowNumber = attrNum(rowEl, "r");
        if (rowNumber !== undefined) row.rowNumber = rowNumber;
        const ht = attrNum(rowEl, "ht");
        if (ht !== undefined) row.height = ht;
        if (String(attr(rowEl, "hidden")) === "1") row.hidden = true;
        if (attr(rowEl, "spans")) row.spans = attr(rowEl, "spans");
        if (String(attr(rowEl, "customFormat")) === "1") row.customFormat = true;
        if (String(attr(rowEl, "thickTop")) === "1") row.thickTop = true;
        if (String(attr(rowEl, "thickBot")) === "1") row.thickBot = true;
        if (String(attr(rowEl, "ph")) === "1") row.ph = true;

        const cells: CellOptions[] = [];
        for (const cellEl of rowEl.elements ?? []) {
          if (cellEl.name !== "c") continue;
          const cell: CellOptions = {};
          const ref = attr(cellEl, "r");
          if (ref) cell.reference = ref;
          const type = attr(cellEl, "t");
          const styleIdx = attrNum(cellEl, "s");
          if (styleIdx !== undefined) {
            // Resolve to a concrete StyleOptions so re-stringify registers it in
            // the fresh Styles table (whose indices may differ). Keep styleIndex
            // as a fallback when the styles table cannot be resolved.
            const resolved =
              ctx && "resolveStyle" in ctx
                ? (ctx as XlsxReadContext).resolveStyle(styleIdx)
                : undefined;
            if (resolved) {
              cell.style = resolved;
            } else {
              cell.styleIndex = styleIdx;
            }
          }

          // Cell value
          const vEl = findChild(cellEl, "v");
          const isEl = findChild(cellEl, "is");

          if (type === "s" && vEl) {
            // Shared string
            const idx = parseInt(textOf(vEl) ?? "", 10);
            cell.value = strings[idx] ?? "";
          } else if (type === "b" && vEl) {
            cell.value = textOf(vEl) === "1";
          } else if (type === "inlineStr" && isEl) {
            const t = findChild(isEl, "t");
            cell.value = textOf(t) ?? "";
          } else if (vEl) {
            const raw = textOf(vEl) ?? "";
            const num = Number(raw);
            cell.value = isNaN(num) ? raw : num;
          }

          // Formula
          const fEl = findChild(cellEl, "f");
          if (fEl) {
            const formula: FormulaOptions = { formula: textOf(fEl) ?? "" };
            const ft = attr(fEl, "t");
            if (ft && ft !== "normal") formula.type = ft as FormulaType;
            const fRef = attr(fEl, "ref");
            if (fRef) formula.reference = fRef;
            const fSi = attrNum(fEl, "si");
            if (fSi !== undefined) formula.sharedIndex = fSi;
            if (String(attr(fEl, "aca")) === "1") formula.aca = true;
            if (String(attr(fEl, "ca")) === "1") formula.ca = true;
            if (String(attr(fEl, "bx")) === "1") formula.bx = true;
            cell.formula = formula;
          }

          cells.push(cell);
        }

        row.cells = cells;
        rows.push(row);
      }
      if (rows.length > 0) result.rows = rows;
    }

    // Row breaks (CT_PageBreak — after sheetData per XSD sequence)
    const rowBreaksEl = findChild(el, "rowBreaks");
    if (rowBreaksEl) {
      const breaks = parsePageBreaks(rowBreaksEl);
      if (breaks.length > 0) result.rowBreaks = breaks;
    }

    // Column breaks (CT_PageBreak)
    const colBreaksEl = findChild(el, "colBreaks");
    if (colBreaksEl) {
      const breaks = parsePageBreaks(colBreaksEl);
      if (breaks.length > 0) result.colBreaks = breaks;
    }

    // Custom properties (CT_CustomProperties — customPr name + r:id)
    const customPropsEl = findChild(el, "customProperties");
    if (customPropsEl) {
      const props: CustomPropertyOptions[] = [];
      for (const cpEl of customPropsEl.elements ?? []) {
        if (cpEl.name !== "customPr") continue;
        const name = attr(cpEl, "name");
        const rId = cpEl.attributes?.["r:id"] as string | undefined;
        if (name !== undefined && rId !== undefined) props.push({ name, rId });
      }
      if (props.length > 0) result.customProperties = props;
    }

    // Cell watches (CT_CellWatches — cellWatch @r)
    const cellWatchesEl = findChild(el, "cellWatches");
    if (cellWatchesEl) {
      const watches: CellWatchOptions[] = [];
      for (const cwEl of cellWatchesEl.elements ?? []) {
        if (cwEl.name !== "cellWatch") continue;
        const r = attr(cwEl, "r");
        if (r) watches.push({ r });
      }
      if (watches.length > 0) result.cellWatches = watches;
    }

    // Legacy drawing header/footer (CT_LegacyDrawing @r:id)
    const legacyHFEl = findChild(el, "legacyDrawingHF");
    if (legacyHFEl) {
      const rId = legacyHFEl.attributes?.["r:id"] as string | undefined;
      if (rId) result.legacyDrawingHF = rId;
    }

    // Data consolidation (CT_DataConsolidate — function/labels/link + dataRefs)
    const dcEl = findChild(el, "dataConsolidate");
    if (dcEl) {
      const dc: DataConsolidateOptions = {};
      const fn = attr(dcEl, "function");
      if (fn) dc.function = fn as DataConsolidateOptions["function"];
      if (String(attr(dcEl, "topLabels")) === "1") dc.topLabels = true;
      if (String(attr(dcEl, "leftLabels")) === "1") dc.leftLabels = true;
      if (String(attr(dcEl, "startLabels")) === "1") dc.startLabels = true;
      if (String(attr(dcEl, "link")) === "1") dc.link = true;
      const refsEl = findChild(dcEl, "dataRefs");
      if (refsEl) {
        const refs: string[] = [];
        for (const refEl of refsEl.elements ?? []) {
          if (refEl.name !== "dataRef") continue;
          const r = attr(refEl, "ref");
          if (r) refs.push(r);
        }
        if (refs.length > 0) dc.refs = refs;
      }
      if (Object.keys(dc).length > 0) result.dataConsolidate = dc;
    }

    // What-if scenarios (CT_Scenarios — current/show + scenario/inputCells)
    const scenariosEl = findChild(el, "scenarios");
    if (scenariosEl) {
      const scenarios: ScenarioDefinition[] = [];
      for (const scEl of scenariosEl.elements ?? []) {
        if (scEl.name !== "scenario") continue;
        const name = attr(scEl, "name");
        if (name === undefined) continue;
        const scenario: ScenarioDefinition = { name, inputCells: [] };
        const count = attrNum(scEl, "count");
        if (count !== undefined) scenario.count = count;
        const user = attr(scEl, "user");
        if (user !== undefined) scenario.user = user;
        const comment = attr(scEl, "comment");
        if (comment !== undefined) scenario.comment = comment;
        if (String(attr(scEl, "hidden")) === "1") scenario.hidden = true;
        if (String(attr(scEl, "locked")) === "1") scenario.locked = true;
        for (const icEl of scEl.elements ?? []) {
          if (icEl.name !== "inputCells") continue;
          const r = attr(icEl, "r");
          const valRaw = attr(icEl, "val");
          if (r === undefined || valRaw === undefined) continue;
          const num = Number(valRaw);
          const cell: ScenarioCellOptions = { r, val: String(num) === valRaw ? num : valRaw };
          if (String(attr(icEl, "deleted")) === "1") cell.deleted = true;
          if (String(attr(icEl, "undone")) === "1") cell.undone = true;
          scenario.inputCells.push(cell);
        }
        scenarios.push(scenario);
      }
      if (scenarios.length > 0) {
        const so: ScenarioOptions = { scenarios };
        const current = attrNum(scenariosEl, "current");
        if (current !== undefined) so.current = current;
        const show = attrNum(scenariosEl, "show");
        if (show !== undefined) so.show = show;
        result.scenarios = so;
      }
    }

    // Custom sheet views (CT_CustomSheetViews — attribute bag per saved view)
    const csvListEl = findChild(el, "customSheetViews");
    if (csvListEl) {
      const views: CustomSheetViewOptions[] = [];
      for (const vEl of csvListEl.elements ?? []) {
        if (vEl.name !== "customSheetView") continue;
        const guid = attr(vEl, "guid");
        if (guid === undefined) continue;
        const view: CustomSheetViewOptions = { guid };
        const scale = attrNum(vEl, "scale");
        if (scale !== undefined) view.scale = scale;
        if (String(attr(vEl, "showPageBreaks")) === "1") view.showPageBreaks = true;
        if (String(attr(vEl, "showFormulas")) === "1") view.showFormulas = true;
        if (String(attr(vEl, "showGridLines")) === "0") view.showGridLines = false;
        if (String(attr(vEl, "showRowCol")) === "0") view.showRowColHeaders = false;
        if (String(attr(vEl, "outlineSymbols")) === "0") view.outlineSymbols = false;
        if (String(attr(vEl, "zeroValues")) === "0") view.zeroValues = false;
        if (String(attr(vEl, "fitToPage")) === "1") view.fitToPage = true;
        if (String(attr(vEl, "printArea")) === "1") view.printArea = true;
        if (String(attr(vEl, "filter")) === "1") view.filter = true;
        if (String(attr(vEl, "showAutoFilter")) === "1") view.showAutoFilter = true;
        if (String(attr(vEl, "hiddenRows")) === "1") view.hiddenRows = true;
        if (String(attr(vEl, "hiddenColumns")) === "1") view.hiddenColumns = true;
        const state = attr(vEl, "state");
        if (state !== undefined) view.state = state as CustomSheetViewOptions["state"];
        if (String(attr(vEl, "filterUnique")) === "1") view.filterUnique = true;
        const viewType = attr(vEl, "view");
        if (viewType !== undefined) view.view = viewType as CustomSheetViewOptions["view"];
        views.push(view);
      }
      if (views.length > 0) result.customSheetViews = views;
    }

    // OLE objects (CT_OleObjects — oleObject attrs + optional objectPr child)
    const oleObjsEl = findChild(el, "oleObjects");
    if (oleObjsEl) {
      const oleObjects: OleObjectOptions[] = [];
      for (const ooEl of oleObjsEl.elements ?? []) {
        if (ooEl.name !== "oleObject") continue;
        const shapeId = attrNum(ooEl, "shapeId");
        if (shapeId === undefined) continue;
        const oo: OleObjectOptions = { shapeId };
        const progId = attr(ooEl, "progId");
        if (progId !== undefined) oo.progId = progId;
        const dvAspect = attr(ooEl, "dvAspect");
        if (dvAspect !== undefined) oo.dvAspect = dvAspect as OleObjectOptions["dvAspect"];
        const link = attr(ooEl, "link");
        if (link !== undefined) oo.link = link;
        const oleUpdate = attr(ooEl, "oleUpdate");
        if (oleUpdate !== undefined) oo.oleUpdate = oleUpdate as OleObjectOptions["oleUpdate"];
        if (String(attr(ooEl, "autoLoad")) === "1") oo.autoLoad = true;
        const ooRid = attr(ooEl, "r:id");
        if (ooRid !== undefined) oo.rId = ooRid;
        const oprEl = findChild(ooEl, "objectPr");
        if (oprEl) {
          const opr: OleObjectPropertiesOptions = {};
          if (String(attr(oprEl, "locked")) === "0") opr.locked = false;
          if (String(attr(oprEl, "defaultSize")) === "0") opr.defaultSize = false;
          if (String(attr(oprEl, "print")) === "0") opr.print = false;
          if (String(attr(oprEl, "disabled")) === "1") opr.disabled = true;
          if (String(attr(oprEl, "uiObject")) === "1") opr.uiObject = true;
          if (String(attr(oprEl, "autoFill")) === "0") opr.autoFill = false;
          if (String(attr(oprEl, "autoLine")) === "0") opr.autoLine = false;
          if (String(attr(oprEl, "autoPict")) === "0") opr.autoPict = false;
          const macro = attr(oprEl, "macro");
          if (macro !== undefined) opr.macro = macro;
          const altText = attr(oprEl, "altText");
          if (altText !== undefined) opr.altText = altText;
          if (String(attr(oprEl, "dde")) === "1") opr.dde = true;
          const oprRid = attr(oprEl, "r:id");
          if (oprRid !== undefined) opr.rId = oprRid;
          if (Object.keys(opr).length > 0) oo.objectPr = opr;
        }
        oleObjects.push(oo);
      }
      if (oleObjects.length > 0) result.oleObjects = oleObjects;
    }

    // Controls (CT_Controls — control attrs + optional controlPr child)
    const controlsEl = findChild(el, "controls");
    if (controlsEl) {
      const controls: ControlOptions[] = [];
      for (const cEl of controlsEl.elements ?? []) {
        if (cEl.name !== "control") continue;
        const shapeId = attrNum(cEl, "shapeId");
        const cRid = attr(cEl, "r:id");
        if (shapeId === undefined || cRid === undefined) continue;
        const c: ControlOptions = { shapeId, rId: cRid };
        const name = attr(cEl, "name");
        if (name !== undefined) c.name = name;
        const prEl = findChild(cEl, "controlPr");
        if (prEl) {
          if (String(attr(prEl, "locked")) === "0") c.locked = false;
          if (String(attr(prEl, "uiObject")) === "1") c.uiObject = true;
          if (String(attr(prEl, "recalcAlways")) === "1") c.recalcAlways = true;
          const linkedCell = attr(prEl, "linkedCell");
          if (linkedCell !== undefined) c.linkedCell = linkedCell;
          const listFillRange = attr(prEl, "listFillRange");
          if (listFillRange !== undefined) c.listFillRange = listFillRange;
          const cf = attr(prEl, "cf");
          if (cf !== undefined) c.cf = cf;
        }
        controls.push(c);
      }
      if (controls.length > 0) result.controls = controls;
    }

    // Web publish items (CT_WebPublishItems — attribute bag per item)
    const wpEl = findChild(el, "webPublishItems");
    if (wpEl) {
      const items: WebPublishItemOptions[] = [];
      for (const wpiEl of wpEl.elements ?? []) {
        if (wpiEl.name !== "webPublishItem") continue;
        const id = attrNum(wpiEl, "id");
        const divId = attr(wpiEl, "divId");
        const sourceType = attr(wpiEl, "sourceType");
        const destinationFile = attr(wpiEl, "destinationFile");
        if (
          id === undefined ||
          divId === undefined ||
          sourceType === undefined ||
          destinationFile === undefined
        )
          continue;
        const wpi: WebPublishItemOptions = {
          id,
          divId,
          sourceType: sourceType as WebPublishItemOptions["sourceType"],
          destinationFile,
        };
        const sourceRef = attr(wpiEl, "sourceRef");
        if (sourceRef !== undefined) wpi.sourceRef = sourceRef;
        const sourceObject = attr(wpiEl, "sourceObject");
        if (sourceObject !== undefined) wpi.sourceObject = sourceObject;
        const title = attr(wpiEl, "title");
        if (title !== undefined) wpi.title = title;
        if (String(attr(wpiEl, "autoRepublish")) === "1") wpi.autoRepublish = true;
        items.push(wpi);
      }
      if (items.length > 0) result.webPublishItems = items;
    }

    // Drawing in header/footer (CT_DrawingHF — r:id + 18 header/footer offsets)
    const drawingHfEl = findChild(el, "drawingHF");
    if (drawingHfEl) {
      const rId = attr(drawingHfEl, "r:id");
      if (rId !== undefined) {
        const dhf: DrawingHfOptions = { rId };
        for (const key of DRAWING_HF_OFFSET_KEYS) {
          const v = attrNum(drawingHfEl, key);
          if (v !== undefined) dhf[key] = v;
        }
        result.drawingHF = dhf;
      }
    }

    // Extension list (CT_ExtensionList — verbatim inner XML, open-ended content)
    const extLstEl = findChild(el, "extLst");
    if (extLstEl) {
      const inner = stringify(extLstEl);
      if (inner) result.ext = inner;
    }

    return result as WorksheetOptions;
  },
};
