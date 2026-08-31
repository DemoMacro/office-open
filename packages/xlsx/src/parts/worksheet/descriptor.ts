/**
 * Worksheet — descriptor for xl/worksheets/sheet{n}.xml.
 *
 * stringify() intentionally throws: the compiler calls stringifyWorksheet()
 * directly (it needs the SharedStrings/Styles accumulators). This descriptor
 * exists for the read path only.
 *
 * @module
 */
import { parseOnOff } from "@office-open/core";
import { xsdConsolidateFunction } from "@office-open/core";
import type { PositiveUniversalMeasure } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrMeasure, attrNum, findChild, stringify, textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { XlsxReadContext } from "../../context";
import { parseAutoFilter, parseSortStateEl } from "../auto-filter";
import { parsePivotArea } from "../pivot-table/parse";
import { parseCfColor, parseCfvo, parsePageBreaks } from "./parse";
import { parseSheetDataRows } from "./sheet-data";
import type {
  CellSmartTagsOptions,
  CellWatchOptions,
  CfColorOptions,
  CfvoOptions,
  ColumnOptions,
  ConditionalFormatOperator,
  ConditionalFormatOptions,
  ConditionalFormatRule,
  ConditionalFormatType,
  ControlOptions,
  AnchorMarkerOptions,
  ObjectAnchorOptions,
  CustomSheetPropertyOptions,
  CustomSheetViewOptions,
  DataConsolidateOptions,
  DataValidationOperator,
  DataValidationOptions,
  DataValidationType,
  DrawingHfOptions,
  FreezePaneOptions,
  HeaderFooterOptions,
  HyperlinkOptions,
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
  PivotSelectionOptions,
  RichTextOptions,
  PrintOptions,
  ProtectedRangeOptions,
  ScenarioCellOptions,
  ScenarioDefinition,
  ScenarioOptions,
  SelectionOptions,
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

    // Resolve shared strings from context (XlsxReadContext). Rich-text
    // entries arrive as RichTextOptions objects and flow into cell.value.
    const strings: (string | RichTextOptions)[] =
      ctx && "sharedStrings" in ctx ? (ctx as XlsxReadContext).sharedStrings : [];

    // Sheet properties
    const sheetPrEl = findChild(el, "sheetPr");
    if (sheetPrEl) {
      const sp: SheetPropertiesOptions = {};
      if (attr(sheetPrEl, "codeName")) sp.codeName = attr(sheetPrEl, "codeName");
      if (parseOnOff(attr(sheetPrEl, "syncHorizontal"))) sp.syncHorizontal = true;
      if (parseOnOff(attr(sheetPrEl, "syncVertical"))) sp.syncVertical = true;
      if (attr(sheetPrEl, "syncRef")) sp.syncRef = attr(sheetPrEl, "syncRef");
      if (parseOnOff(attr(sheetPrEl, "transitionEvaluation"))) sp.transitionEvaluation = true;
      if (parseOnOff(attr(sheetPrEl, "transitionEntry"))) sp.transitionEntry = true;
      // XSD defaults true — only the explicit "0" carries information back.
      if (String(attr(sheetPrEl, "published")) === "0") sp.published = false;
      if (parseOnOff(attr(sheetPrEl, "filterMode"))) sp.filterMode = true;
      if (String(attr(sheetPrEl, "enableFormatConditionsCalculation")) === "0")
        sp.enableFormatConditionsCalculation = false;

      const outlinePr = findChild(sheetPrEl, "outlinePr");
      if (outlinePr) {
        if (parseOnOff(attr(outlinePr, "applyStyles"))) sp.outlineApplyStyles = true;
        if (String(attr(outlinePr, "showOutlineSymbols")) === "0") sp.outlineShowSymbols = false;
        if (String(attr(outlinePr, "summaryBelow")) === "0") sp.outlineSummaryBelow = false;
        if (String(attr(outlinePr, "summaryRight")) === "0") sp.outlineSummaryRight = false;
      }

      // pageSetUpPr (inside sheetPr) — stash on result.pageSetup; merged into
      // the <pageSetup> parse below, which owns result.pageSetup.
      const pageSetUpPr = findChild(sheetPrEl, "pageSetUpPr");
      if (pageSetUpPr) {
        const psup: Partial<PageSetupOptions> = {};
        // Present-as-written: autoPageBreaks defaults true, so an explicit "0"
        // carries information and must survive the round-trip.
        if (attr(pageSetUpPr, "fitToPage") !== undefined)
          psup.fitToPage = parseOnOff(attr(pageSetUpPr, "fitToPage")) ?? false;
        if (attr(pageSetUpPr, "autoPageBreaks") !== undefined)
          psup.autoPageBreaks = parseOnOff(attr(pageSetUpPr, "autoPageBreaks")) ?? true;
        if (Object.keys(psup).length > 0) pageSetUpPrCache = psup;
      }
      if (Object.keys(sp).length > 0) result.properties = sp;

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
          sv.tabSelected = parseOnOff(attr(svEl, "tabSelected")) ?? true;
        if (parseOnOff(attr(svEl, "rightToLeft"))) sv.rightToLeft = true;
        if (parseOnOff(attr(svEl, "windowProtection"))) sv.windowProtection = true;
        if (parseOnOff(attr(svEl, "showFormulas"))) sv.showFormulas = true;
        if (String(attr(svEl, "showRuler")) === "0") sv.showRuler = false;
        if (String(attr(svEl, "showOutlineSymbols")) === "0") sv.showOutlineSymbols = false;
        if (String(attr(svEl, "defaultGridColor")) === "0") sv.defaultGridColor = false;
        if (String(attr(svEl, "showWhiteSpace")) === "0") sv.showWhiteSpace = false;
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

        // Freeze/split pane — CT_Pane (any state; frozen is the authoring default)
        const paneEl = findChild(svEl, "pane");
        if (paneEl) {
          const fp: FreezePaneOptions = {};
          const ys = attrNum(paneEl, "ySplit");
          if (ys && ys > 0) fp.row = ys;
          const xs = attrNum(paneEl, "xSplit");
          if (xs && xs > 0) fp.col = xs;
          if (attr(paneEl, "state") === "split") fp.split = true;
          const paneTopLeft = attr(paneEl, "topLeftCell");
          if (paneTopLeft) fp.topLeftCell = paneTopLeft;
          const activePane = attr(paneEl, "activePane");
          if (activePane) fp.activePane = activePane as FreezePaneOptions["activePane"];
          if (Object.keys(fp).length > 0) result.freezePanes = fp;
        }

        // Selections (CT_Selection — one per pane)
        const selections: SelectionOptions[] = [];
        for (const selEl of svEl.elements ?? []) {
          if (selEl.name !== "selection") continue;
          const sel: SelectionOptions = {};
          const pane = attr(selEl, "pane");
          if (pane) sel.pane = pane as SelectionOptions["pane"];
          const activeCell = attr(selEl, "activeCell");
          if (activeCell) sel.activeCell = activeCell;
          const acId = attrNum(selEl, "activeCellId");
          if (acId !== undefined) sel.activeCellId = acId;
          const sqref = attr(selEl, "sqref");
          if (sqref) sel.sqref = sqref;
          selections.push(sel);
        }
        if (selections.length > 0) result.selection = selections;

        // Pivot selection (CT_PivotSelection)
        const psEl = findChild(svEl, "pivotSelection");
        if (psEl) {
          const ps: Partial<PivotSelectionOptions> = {};
          const pane = attr(psEl, "pane");
          if (pane) ps.pane = pane as PivotSelectionOptions["pane"];
          if (parseOnOff(attr(psEl, "showHeader"))) ps.showHeader = true;
          if (parseOnOff(attr(psEl, "label"))) ps.label = true;
          if (parseOnOff(attr(psEl, "data"))) ps.data = true;
          if (parseOnOff(attr(psEl, "extendable"))) ps.extendable = true;
          const count = attrNum(psEl, "count");
          if (count !== undefined) ps.count = count;
          const axis = attr(psEl, "axis");
          if (axis) ps.axis = axis as PivotSelectionOptions["axis"];
          for (const key of [
            "dimension",
            "start",
            "min",
            "max",
            "activeRow",
            "activeCol",
            "previousRow",
            "previousCol",
            "click",
          ] as const) {
            const v = attrNum(psEl, key);
            if (v !== undefined) ps[key] = v;
          }
          const rId = attr(psEl, "r:id");
          if (rId) ps.rId = rId;
          const paEl = findChild(psEl, "pivotArea");
          if (paEl) ps.pivotArea = parsePivotArea(paEl) as PivotSelectionOptions["pivotArea"];
          result.pivotSelection = ps as PivotSelectionOptions;
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
      if (parseOnOff(attr(sfpEl, "zeroHeight"))) sfp.zeroHeight = true;
      if (parseOnOff(attr(sfpEl, "thickTop"))) sfp.thickTop = true;
      if (parseOnOff(attr(sfpEl, "thickBottom"))) sfp.thickBottom = true;
      const olr = attrNum(sfpEl, "outlineLevelRow");
      if (olr !== undefined) sfp.outlineLevelRow = olr;
      const olc = attrNum(sfpEl, "outlineLevelCol");
      if (olc !== undefined) sfp.outlineLevelCol = olc;
      result.sheetFormat = sfp;
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
        if (parseOnOff(attr(colEl, "hidden"))) col.hidden = true;
        if (parseOnOff(attr(colEl, "customWidth"))) col.customWidth = true;
        const ol = attrNum(colEl, "outlineLevel");
        if (ol !== undefined) col.outlineLevel = ol;
        if (parseOnOff(attr(colEl, "collapsed"))) col.collapsed = true;
        if (parseOnOff(attr(colEl, "bestFit"))) col.bestFit = true;
        if (parseOnOff(attr(colEl, "phonetic"))) col.phonetic = true;
        columns.push(col);
      }
      if (columns.length > 0) result.columns = columns;
    }

    // Sheet protection
    const protEl = findChild(el, "sheetProtection");
    if (protEl?.attributes) {
      result.protection = parseSheetProtectionEl(protEl);
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
        // @password (legacy hash) not read back — see parseSheetProtectionEl note.
        if (attr(rEl, "algorithmName")) r.algorithmName = attr(rEl, "algorithmName");
        if (attr(rEl, "hashValue")) r.hashValue = attr(rEl, "hashValue");
        if (attr(rEl, "saltValue")) r.saltValue = attr(rEl, "saltValue");
        const spinCount = attrNum(rEl, "spinCount");
        if (spinCount !== undefined) r.spinCount = spinCount;
        const sdAttr = attr(rEl, "securityDescriptor");
        if (sdAttr !== undefined) r.securityDescriptor = sdAttr;
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

    // Sheet-level sort state (after autoFilter per XSD sequence)
    const ssEl = findChild(el, "sortState");
    if (ssEl) {
      result.sortState = parseSortStateEl(ssEl);
    }

    // Merge cells
    const mcEl = findChild(el, "mergeCells");
    if (mcEl) {
      const merges: MergeCellOptions[] = [];
      for (const mEl of mcEl.elements ?? []) {
        if (mEl.name !== "mergeCell") continue;
        const ref = attr(mEl, "ref") ?? "";
        if (ref) merges.push({ ref });
      }
      if (merges.length > 0) result.mergeCells = merges;
    }

    // Conditional formatting
    const cfEls = el.elements?.filter((e) => e.name === "conditionalFormatting") ?? [];
    if (cfEls.length > 0) {
      const cfs: ConditionalFormatOptions[] = [];
      for (const cfEl of cfEls) {
        const sqref = attr(cfEl, "sqref") ?? "";
        const pivot = parseOnOff(attr(cfEl, "pivot")) ?? false;
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
          if (parseOnOff(attr(ruleEl, "stopIfTrue"))) rule.stopIfTrue = true;
          const tpVal = attr(ruleEl, "timePeriod");
          if (tpVal) rule.timePeriod = tpVal as ConditionalFormatRule["timePeriod"];
          const rank = attrNum(ruleEl, "rank");
          if (rank !== undefined) rule.rank = rank;
          if (parseOnOff(attr(ruleEl, "bottom"))) rule.bottom = true;
          if (parseOnOff(attr(ruleEl, "percent"))) rule.percent = true;
          const textVal = attr(ruleEl, "text");
          if (textVal) rule.text = textVal;
          if (parseOnOff(attr(ruleEl, "equalAverage"))) rule.equalAverage = true;
          if (String(attr(ruleEl, "aboveAverage")) === "0") rule.aboveAverage = false;
          const stdDev = attrNum(ruleEl, "stdDev");
          if (stdDev !== undefined) rule.stdDev = stdDev;

          // Color scale
          const csEl = findChild(ruleEl, "colorScale");
          if (csEl) {
            const cfvo: CfvoOptions[] = [];
            const colors: CfColorOptions[] = [];
            for (const child of csEl.elements ?? []) {
              if (child.name === "cfvo") cfvo.push(parseCfvo(child));
              if (child.name === "color") {
                const c = parseCfColor(child);
                if (c) colors.push(c);
              }
            }
            rule.colorScale = { cfvo, colors };
          }

          // Data bar
          const dbEl = findChild(ruleEl, "dataBar");
          if (dbEl) {
            const cfvo: CfvoOptions[] = [];
            let color: CfColorOptions | undefined;
            for (const child of dbEl.elements ?? []) {
              if (child.name === "cfvo") cfvo.push(parseCfvo(child));
              if (child.name === "color") color = parseCfColor(child) ?? { indexed: 0 };
            }
            rule.dataBar = {
              cfvo: cfvo as [CfvoOptions, CfvoOptions],
              color: color ?? { indexed: 0 },
            };
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
            if (parseOnOff(attr(isEl, "reverse"))) iconSet.reverse = true;
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
        cfs.push(pivot ? { sqref, pivot, rules } : { sqref, rules });
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
        if (parseOnOff(attr(dEl, "allowBlank"))) dv.allowBlank = true;
        if (parseOnOff(attr(dEl, "showErrorMessage"))) dv.showErrorMessage = true;
        if (parseOnOff(attr(dEl, "showInputMessage"))) dv.showInputMessage = true;
        if (attr(dEl, "errorTitle")) dv.errorTitle = attr(dEl, "errorTitle");
        if (attr(dEl, "error")) dv.error = attr(dEl, "error");
        if (attr(dEl, "promptTitle")) dv.promptTitle = attr(dEl, "promptTitle");
        if (attr(dEl, "prompt")) dv.prompt = attr(dEl, "prompt");
        const esVal = attr(dEl, "errorStyle");
        if (esVal) dv.errorStyle = esVal as DataValidationOptions["errorStyle"];
        const imVal = attr(dEl, "imeMode");
        if (imVal) dv.imeMode = imVal as DataValidationOptions["imeMode"];
        if (parseOnOff(attr(dEl, "showDropDown"))) dv.showDropDown = true;

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
        const hl: HyperlinkOptions = { cell: attr(hEl, "ref") ?? "" };
        // r:id starts as the raw relationship id; the parse pipeline resolves
        // it to the real target URL afterwards.
        if (rId) hl.url = rId;
        const location = attr(hEl, "location");
        if (location) hl.location = location;
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
      if (parseOnOff(attr(poEl, "horizontalCentered"))) po.horizontalCentered = true;
      if (parseOnOff(attr(poEl, "verticalCentered"))) po.verticalCentered = true;
      if (parseOnOff(attr(poEl, "headings"))) po.headings = true;
      if (parseOnOff(attr(poEl, "gridLines"))) po.gridLines = true;
      if (String(attr(poEl, "gridLinesSet")) === "0") po.gridLinesSet = false;
      result.printOptions = po;
    }

    // Page setup
    const psEl = findChild(el, "pageSetup");
    if (psEl) {
      result.pageSetup = parsePageSetupEl(psEl, pageSetUpPrCache);
    } else if (pageSetUpPrCache) {
      result.pageSetup = pageSetUpPrCache;
    }

    // Header/footer
    const hfEl = findChild(el, "headerFooter");
    if (hfEl) {
      result.headerFooter = parseHeaderFooterEl(hfEl);
    }

    // Ignored errors
    const ieEl = findChild(el, "ignoredErrors");
    if (ieEl) {
      const errors: IgnoredErrorOptions[] = [];
      for (const eEl of ieEl.elements ?? []) {
        if (eEl.name !== "ignoredError") continue;
        const ie: IgnoredErrorOptions = { sqref: attr(eEl, "sqref") ?? "" };
        if (parseOnOff(attr(eEl, "evalError"))) ie.evalError = true;
        if (parseOnOff(attr(eEl, "twoDigitTextYear"))) ie.twoDigitTextYear = true;
        if (parseOnOff(attr(eEl, "numberStoredAsText"))) ie.numberStoredAsText = true;
        if (parseOnOff(attr(eEl, "formula"))) ie.formula = true;
        if (parseOnOff(attr(eEl, "formulaRange"))) ie.formulaRange = true;
        if (parseOnOff(attr(eEl, "unlockedFormula"))) ie.unlockedFormula = true;
        if (parseOnOff(attr(eEl, "emptyCellReference"))) ie.emptyCellReference = true;
        if (parseOnOff(attr(eEl, "listDataValidation"))) ie.listDataValidation = true;
        if (parseOnOff(attr(eEl, "calculatedColumn"))) ie.calculatedColumn = true;
        errors.push(ie);
      }
      result.ignoredErrors = errors;
    }

    // Cell smart tags (CT_SmartTags)
    const smartTagsEl = findChild(el, "smartTags");
    if (smartTagsEl) {
      const entries: CellSmartTagsOptions[] = [];
      for (const cstEl of smartTagsEl.elements ?? []) {
        if (cstEl.name !== "cellSmartTags") continue;
        const cst: CellSmartTagsOptions = {
          reference: attr(cstEl, "r") ?? "",
          smartTags: [],
        };
        for (const stEl of cstEl.elements ?? []) {
          if (stEl.name !== "cellSmartTag") continue;
          const st: {
            type: number;
            deleted?: boolean;
            xmlBased?: boolean;
            properties?: { key: string; val: string }[];
          } = {
            type: attrNum(stEl, "type") ?? 0,
          };
          if (parseOnOff(attr(stEl, "deleted"))) st.deleted = true;
          if (parseOnOff(attr(stEl, "xmlBased"))) st.xmlBased = true;
          const prEls = stEl.elements?.filter((e) => e.name === "cellSmartTagPr") ?? [];
          if (prEls.length > 0) {
            st.properties = prEls.map((prEl) => ({
              key: attr(prEl, "key") ?? "",
              val: attr(prEl, "val") ?? "",
            }));
          }
          cst.smartTags.push(st);
        }
        entries.push(cst);
      }
      if (entries.length > 0) result.smartTags = entries;
    }

    // Phonetic properties
    const ppEl = findChild(el, "phoneticPr");
    if (ppEl) {
      const pp: PhoneticPropertiesOptions = { fontId: attrNum(ppEl, "fontId") ?? 0 };
      const ppType = attr(ppEl, "type");
      if (ppType) pp.type = ppType as PhoneticPropertiesOptions["type"];
      const ppAlign = attr(ppEl, "alignment");
      if (ppAlign) pp.alignment = ppAlign as PhoneticPropertiesOptions["alignment"];
      result.phonetic = pp;
    }

    // Sheet calc properties
    const scEl = findChild(el, "sheetCalcPr");
    if (scEl) {
      const sc: SheetCalculationPropertiesOptions = {};
      if (parseOnOff(attr(scEl, "fullCalcOnLoad"))) sc.fullCalcOnLoad = true;
      result.calculation = sc;
    }

    // Sheet data (rows and cells) — parsed by the dedicated sheetData scanner.
    // The archive read path defers sheetData (Element.raw, children never
    // materialized); a hand-built or non-deferred tree normalizes through
    // stringify first, so there is exactly one row parsing implementation.
    const sheetDataEl = findChild(el, "sheetData");
    if (sheetDataEl) {
      const rows = parseSheetDataRows(sheetDataEl.raw ?? stringify(sheetDataEl), strings);
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
      const props: CustomSheetPropertyOptions[] = [];
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
        if (r) watches.push({ reference: r });
      }
      if (watches.length > 0) result.cellWatches = watches;
    }

    // Legacy drawing header/footer (CT_LegacyDrawing @r:id)
    const legacyHFEl = findChild(el, "legacyDrawingHF");
    if (legacyHFEl) {
      const rId = legacyHFEl.attributes?.["r:id"] as string | undefined;
      if (rId) result.legacyDrawingHF = rId;
    }

    // Round-trip drawing/legacyDrawing references: the referenced part may
    // pass through verbatim (anchors the bridge does not map onto options),
    // so keep the original reference id to re-emit the element.
    const drawingRefEl = findChild(el, "drawing");
    if (drawingRefEl) {
      const rId = drawingRefEl.attributes?.["r:id"] as string | undefined;
      if (rId) result.drawingRid = rId;
    }
    const legacyRefEl = findChild(el, "legacyDrawing");
    if (legacyRefEl) {
      const rId = legacyRefEl.attributes?.["r:id"] as string | undefined;
      if (rId) result.legacyDrawingRid = rId;
    }

    // Data consolidation (CT_DataConsolidate — function/labels/link + dataRefs)
    const dcEl = findChild(el, "dataConsolidate");
    if (dcEl) {
      const dc: DataConsolidateOptions = {};
      const fn = attr(dcEl, "function");
      if (fn) dc.function = xsdConsolidateFunction.from(fn) as DataConsolidateOptions["function"];
      if (parseOnOff(attr(dcEl, "topLabels"))) dc.topLabels = true;
      if (parseOnOff(attr(dcEl, "leftLabels"))) dc.leftLabels = true;
      if (parseOnOff(attr(dcEl, "startLabels"))) dc.startLabels = true;
      if (parseOnOff(attr(dcEl, "link"))) dc.link = true;
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
        if (parseOnOff(attr(scEl, "hidden"))) scenario.hidden = true;
        if (parseOnOff(attr(scEl, "locked"))) scenario.locked = true;
        for (const icEl of scEl.elements ?? []) {
          if (icEl.name !== "inputCells") continue;
          const r = attr(icEl, "r");
          const valRaw = attr(icEl, "val");
          if (r === undefined || valRaw === undefined) continue;
          const num = Number(valRaw);
          const cell: ScenarioCellOptions = {
            reference: r,
            val: String(num) === valRaw ? num : valRaw,
          };
          if (parseOnOff(attr(icEl, "deleted"))) cell.deleted = true;
          if (parseOnOff(attr(icEl, "undone"))) cell.undone = true;
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
        if (parseOnOff(attr(vEl, "showPageBreaks"))) view.showPageBreaks = true;
        if (parseOnOff(attr(vEl, "showFormulas"))) view.showFormulas = true;
        if (String(attr(vEl, "showGridLines")) === "0") view.showGridLines = false;
        if (String(attr(vEl, "showRowCol")) === "0") view.showRowColHeaders = false;
        if (String(attr(vEl, "outlineSymbols")) === "0") view.outlineSymbols = false;
        if (String(attr(vEl, "zeroValues")) === "0") view.zeroValues = false;
        if (parseOnOff(attr(vEl, "fitToPage"))) view.fitToPage = true;
        if (parseOnOff(attr(vEl, "printArea"))) view.printArea = true;
        if (parseOnOff(attr(vEl, "filter"))) view.filter = true;
        if (parseOnOff(attr(vEl, "showAutoFilter"))) view.showAutoFilter = true;
        if (parseOnOff(attr(vEl, "hiddenRows"))) view.hiddenRows = true;
        if (parseOnOff(attr(vEl, "hiddenColumns"))) view.hiddenColumns = true;
        const state = attr(vEl, "state");
        if (state !== undefined) view.state = state as CustomSheetViewOptions["state"];
        if (parseOnOff(attr(vEl, "filterUnique"))) view.filterUnique = true;
        const viewType = attr(vEl, "view");
        if (viewType !== undefined) view.view = viewType as CustomSheetViewOptions["view"];
        views.push(view);
      }
      if (views.length > 0) result.customSheetViews = views;
    }

    /**
     * Unwrap mc:AlternateContent around an oleObject/control: the mc:Choice leg
     * carries the full element, mc:Fallback the bare one. Returns the element
     * found in either the Choice or directly (unwrapped legacy form), plus
     * whether a wrapper was present.
     */
    function unwrapAlternateContent(
      el: Element,
      innerTag: string,
    ): { element: Element | undefined; wrapped: boolean } {
      if (el.name !== "mc:AlternateContent") {
        return { element: el.name === innerTag ? el : undefined, wrapped: false };
      }
      const choice = (el.elements ?? []).find((c) => c.name === "mc:Choice");
      const inner = (choice?.elements ?? []).find((c) => c.name === innerTag);
      return { element: inner, wrapped: inner !== undefined };
    }

    /** Read the anchor element inside objectPr/controlPr (from/to CT_Marker corners). */
    function readEmbeddedAnchor(anchorEl: Element): ObjectAnchorOptions | undefined {
      const fromEl = findChild(anchorEl, "from");
      const toEl = findChild(anchorEl, "to");
      if (!fromEl || !toEl) return undefined;
      // col/colOff/row/rowOff are xdr-prefixed child elements (CT_Marker), not attrs.
      const markerNum = (m: Element, tag: string): number | undefined => {
        const child = findChild(m, tag);
        const n = child === undefined ? undefined : Number(textOf(child));
        return n === undefined || Number.isNaN(n) ? undefined : n;
      };
      const marker = (m: Element): AnchorMarkerOptions | undefined => {
        const col = markerNum(m, "xdr:col");
        const row = markerNum(m, "xdr:row");
        if (col === undefined || row === undefined) return undefined;
        return {
          col,
          row,
          ...(markerNum(m, "xdr:colOff") !== undefined
            ? { colOff: markerNum(m, "xdr:colOff")! }
            : {}),
          ...(markerNum(m, "xdr:rowOff") !== undefined
            ? { rowOff: markerNum(m, "xdr:rowOff")! }
            : {}),
        };
      };
      const from = marker(fromEl);
      const to = marker(toEl);
      if (!from || !to) return undefined;
      return {
        from,
        to,
        ...(parseOnOff(attr(anchorEl, "moveWithCells")) ? { moveWithCells: true } : {}),
        ...(parseOnOff(attr(anchorEl, "sizeWithCells")) ? { sizeWithCells: true } : {}),
      };
    }

    // OLE objects (CT_OleObjects — oleObject attrs + optional objectPr child)
    const oleObjsEl = findChild(el, "oleObjects");
    if (oleObjsEl) {
      const oleObjects: OleObjectOptions[] = [];
      for (const rawEl of oleObjsEl.elements ?? []) {
        // Excel 2010+ wraps each oleObject in mc:AlternateContent: the Choice
        // carries the full element (objectPr + anchor), the Fallback the bare
        // one. Parse the Choice leg and remember the wrapper for re-emission.
        const unwrapped = unwrapAlternateContent(rawEl, "oleObject");
        const ooEl = unwrapped.element;
        if (!ooEl) continue;
        const shapeId = attrNum(ooEl, "shapeId");
        if (shapeId === undefined) continue;
        const oo: OleObjectOptions = { shapeId };
        if (unwrapped.wrapped) oo.alternateContent = true;
        const progId = attr(ooEl, "progId");
        if (progId !== undefined) oo.progId = progId;
        const dvAspect = attr(ooEl, "dvAspect");
        if (dvAspect !== undefined) oo.dvAspect = dvAspect as OleObjectOptions["dvAspect"];
        const link = attr(ooEl, "link");
        if (link !== undefined) oo.link = link;
        const oleUpdate = attr(ooEl, "oleUpdate");
        if (oleUpdate !== undefined) oo.oleUpdate = oleUpdate as OleObjectOptions["oleUpdate"];
        if (parseOnOff(attr(ooEl, "autoLoad"))) oo.autoLoad = true;
        const ooRid = attr(ooEl, "r:id");
        if (ooRid !== undefined) oo.rId = ooRid;
        const oprEl = findChild(ooEl, "objectPr");
        if (oprEl) {
          const opr: OleObjectPropertiesOptions = {};
          if (String(attr(oprEl, "locked")) === "0") opr.locked = false;
          if (String(attr(oprEl, "defaultSize")) === "0") opr.defaultSize = false;
          if (String(attr(oprEl, "print")) === "0") opr.print = false;
          if (parseOnOff(attr(oprEl, "disabled"))) opr.disabled = true;
          if (parseOnOff(attr(oprEl, "uiObject"))) opr.uiObject = true;
          if (String(attr(oprEl, "autoFill")) === "0") opr.autoFill = false;
          if (String(attr(oprEl, "autoLine")) === "0") opr.autoLine = false;
          if (String(attr(oprEl, "autoPict")) === "0") opr.autoPict = false;
          const macro = attr(oprEl, "macro");
          if (macro !== undefined) opr.macro = macro;
          const altText = attr(oprEl, "altText");
          if (altText !== undefined) opr.altText = altText;
          if (parseOnOff(attr(oprEl, "dde"))) opr.dde = true;
          const oprRid = attr(oprEl, "r:id");
          if (oprRid !== undefined) opr.iconRid = oprRid;
          const anchorEl = findChild(oprEl, "anchor");
          if (anchorEl) {
            const anchor = readEmbeddedAnchor(anchorEl);
            if (anchor) opr.anchor = anchor;
          }
          if (Object.keys(opr).length > 0) oo.properties = opr;
        }
        oleObjects.push(oo);
      }
      if (oleObjects.length > 0) result.oleObjects = oleObjects;
    }

    // Controls (CT_Controls — control attrs + optional controlPr child)
    const controlsEl = findChild(el, "controls");
    if (controlsEl) {
      const controls: ControlOptions[] = [];
      for (const rawEl of controlsEl.elements ?? []) {
        // Same mc:AlternateContent wrapper as oleObjects (Excel 2010+ form).
        const unwrapped = unwrapAlternateContent(rawEl, "control");
        const cEl = unwrapped.element;
        if (!cEl) continue;
        const shapeId = attrNum(cEl, "shapeId");
        const cRid = attr(cEl, "r:id");
        if (shapeId === undefined || cRid === undefined) continue;
        const c: ControlOptions = { shapeId, rId: cRid };
        if (unwrapped.wrapped) c.alternateContent = true;
        const name = attr(cEl, "name");
        if (name !== undefined) c.name = name;
        const prEl = findChild(cEl, "controlPr");
        if (prEl) {
          if (String(attr(prEl, "locked")) === "0") c.locked = false;
          if (parseOnOff(attr(prEl, "uiObject"))) c.uiObject = true;
          if (parseOnOff(attr(prEl, "recalcAlways"))) c.recalcAlways = true;
          const linkedCell = attr(prEl, "linkedCell");
          if (linkedCell !== undefined) c.linkedCell = linkedCell;
          const listFillRange = attr(prEl, "listFillRange");
          if (listFillRange !== undefined) c.listFillRange = listFillRange;
          const cf = attr(prEl, "cf");
          if (cf !== undefined) c.formula = cf;
          if (String(attr(prEl, "defaultSize")) === "0") c.defaultSize = false;
          if (String(attr(prEl, "autoLine")) === "0") c.autoLine = false;
          if (String(attr(prEl, "autoPict")) === "0") c.autoPict = false;
          const prRid = attr(prEl, "r:id");
          if (prRid !== undefined) c.iconRid = prRid;
          const anchorEl = findChild(prEl, "anchor");
          if (anchorEl) {
            const anchor = readEmbeddedAnchor(anchorEl);
            if (anchor) c.anchor = anchor;
          }
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
        if (parseOnOff(attr(wpiEl, "autoRepublish"))) wpi.autoRepublish = true;
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

// ── Shared page-setup parse helpers (worksheet / dialogsheet / chartsheet) ──

/**
 * Parse a CT_PageSetup element. `pageSetUpPrCache` carries the sheetPr-level
 * CT_PageSetUpPr flags (fitToPage/autoPageBreaks) parsed earlier so they merge
 * into the same PageSetupOptions object.
 */
export function parsePageSetupEl(
  el: Element,
  pageSetUpPrCache?: Partial<PageSetupOptions>,
): PageSetupOptions {
  const ps: PageSetupOptions = {};
  const pz = attrNum(el, "paperSize");
  if (pz !== undefined) ps.paperSize = pz;
  const ph = attrMeasure(el, "paperHeight");
  if (ph !== undefined) ps.paperHeight = ph as number | PositiveUniversalMeasure;
  const pw = attrMeasure(el, "paperWidth");
  if (pw !== undefined) ps.paperWidth = pw as number | PositiveUniversalMeasure;
  const orientVal = attr(el, "orientation");
  if (orientVal) ps.orientation = orientVal as PageOrientation;
  const sc = attrNum(el, "scale");
  if (sc !== undefined) ps.scale = sc;
  const ftw = attrNum(el, "fitToWidth");
  if (ftw !== undefined) ps.fitToWidth = ftw;
  const fth = attrNum(el, "fitToHeight");
  if (fth !== undefined) ps.fitToHeight = fth;
  const pageOrderVal = attr(el, "pageOrder");
  if (pageOrderVal) ps.pageOrder = pageOrderVal as PageSetupOptions["pageOrder"];
  if (parseOnOff(attr(el, "useFirstPageNumber"))) ps.useFirstPageNumber = true;
  const fpn = attrNum(el, "firstPageNumber");
  if (fpn !== undefined) ps.firstPageNumber = fpn;
  // XSD default true — only the explicit "0" carries information back.
  if (String(attr(el, "usePrinterDefaults")) === "0") ps.usePrinterDefaults = false;
  if (parseOnOff(attr(el, "blackAndWhite"))) ps.blackAndWhite = true;
  if (parseOnOff(attr(el, "draft"))) ps.draft = true;
  const cc = attr(el, "cellComments");
  if (cc) ps.cellComments = cc as PageSetupOptions["cellComments"];
  const err = attr(el, "errors");
  if (err) ps.errors = err as PageSetupOptions["errors"];
  const hdpi = attrNum(el, "horizontalDpi");
  if (hdpi !== undefined) ps.horizontalDpi = hdpi;
  const vdpi = attrNum(el, "verticalDpi");
  if (vdpi !== undefined) ps.verticalDpi = vdpi;
  const copies = attrNum(el, "copies");
  if (copies !== undefined) ps.copies = copies;
  const psRid = attr(el, "r:id");
  if (psRid) ps.printerSettingsRId = psRid;
  if (pageSetUpPrCache) Object.assign(ps, pageSetUpPrCache);
  return ps;
}

/** Parse a CT_HeaderFooter element. */
export function parseHeaderFooterEl(el: Element): HeaderFooterOptions {
  const hf: HeaderFooterOptions = {};
  if (parseOnOff(attr(el, "differentOddEven"))) hf.differentOddEven = true;
  if (parseOnOff(attr(el, "differentFirst"))) hf.differentFirst = true;
  if (String(attr(el, "scaleWithDoc")) === "0") hf.scaleWithDoc = false;
  if (String(attr(el, "alignWithMargins")) === "0") hf.alignWithMargins = false;
  const oh = findChild(el, "oddHeader");
  if (oh) hf.oddHeader = textOf(oh);
  const of2 = findChild(el, "oddFooter");
  if (of2) hf.oddFooter = textOf(of2);
  const eh = findChild(el, "evenHeader");
  if (eh) hf.evenHeader = textOf(eh);
  const ef = findChild(el, "evenFooter");
  if (ef) hf.evenFooter = textOf(ef);
  const fh = findChild(el, "firstHeader");
  if (fh) hf.firstHeader = textOf(fh);
  const ff = findChild(el, "firstFooter");
  if (ff) hf.firstFooter = textOf(ff);
  return hf;
}

/** Parse a CT_SheetProtection element. */
export function parseSheetProtectionEl(el: Element): SheetProtectionOptions {
  const prot: SheetProtectionOptions = {};
  // NOTE: @password (legacy hash) is deliberately not read back — the
  // password field is plaintext authoring input and stringify hashes it,
  // so carrying the hash would double-hash on round-trip. The modern
  // algorithmName/hashValue/saltValue/spinCount quadruplet round-trips as-is.
  if (attr(el, "algorithmName")) prot.algorithmName = attr(el, "algorithmName");
  if (attr(el, "hashValue")) prot.hashValue = attr(el, "hashValue");
  if (attr(el, "saltValue")) prot.saltValue = attr(el, "saltValue");
  const spin = attrNum(el, "spinCount");
  if (spin !== undefined) prot.spinCount = spin;
  if (parseOnOff(attr(el, "sheet"))) prot.sheet = true;
  if (parseOnOff(attr(el, "objects"))) prot.objects = true;
  if (parseOnOff(attr(el, "scenarios"))) prot.scenarios = true;
  if (String(attr(el, "formatCells")) === "0") prot.formatCells = false;
  if (String(attr(el, "formatColumns")) === "0") prot.formatColumns = false;
  if (String(attr(el, "formatRows")) === "0") prot.formatRows = false;
  if (String(attr(el, "insertColumns")) === "0") prot.insertColumns = false;
  if (String(attr(el, "insertRows")) === "0") prot.insertRows = false;
  if (String(attr(el, "insertHyperlinks")) === "0") prot.insertHyperlinks = false;
  if (String(attr(el, "deleteColumns")) === "0") prot.deleteColumns = false;
  if (String(attr(el, "deleteRows")) === "0") prot.deleteRows = false;
  if (parseOnOff(attr(el, "selectLockedCells"))) prot.selectLockedCells = true;
  if (String(attr(el, "sort")) === "0") prot.sort = false;
  if (String(attr(el, "autoFilter")) === "0") prot.autoFilter = false;
  if (String(attr(el, "pivotTables")) === "0") prot.pivotTables = false;
  if (parseOnOff(attr(el, "selectUnlockedCells"))) prot.selectUnlockedCells = true;
  return prot;
}
