/**
 * Worksheet — stringify implementation for xl/worksheets/sheet{n}.xml.
 *
 * Zero-allocation fast path: directly concatenates XML string, bypassing
 * the intermediate object tree entirely.
 *
 * @module
 */
import {
  convertToInch,
  convertToPt,
  derivePasswordHash,
  type UniversalMeasure,
} from "@office-open/core";
import { xsdConsolidateFunction } from "@office-open/core";
import { attrs, escapeXml, selfCloseElement } from "@office-open/xml";
import { columnToLetter, dateToSerialNumber, hashPassword } from "@util/index";

import { colorAttributes } from "../../shared/color";
import type { ColorOptions } from "../../shared/color";
import { stringifyAutoFilter, stringifySortStateXml } from "../auto-filter";
import { buildPivotAreaXml } from "../pivot-table/stringify";
import { buildRstXml, tElement } from "../shared-strings";
import type { SharedStrings } from "../shared-strings";
import type { Styles } from "../styles";
import { FormulaType } from "./types";
import type { AnchorMarkerOptions, ObjectAnchorOptions } from "./types";
import type {
  CellOptions,
  CfvoOptions,
  FormulaOptions,
  HeaderFooterOptions,
  HeaderFooterPartName,
  PageBreakOptions,
  PageSetupOptions,
  PivotSelectionOptions,
  PrintOptions,
  RowOptions,
  SelectionOptions,
  SheetProtectionOptions,
  SheetViewOptions,
  WorksheetContext,
  WorksheetOptions,
} from "./types";

/** CT_Break list under rowBreaks/colBreaks — identical shape, only the tag differs. */
export function stringifyPageBreaksXml(
  tag: "rowBreaks" | "colBreaks",
  list: PageBreakOptions[],
): string {
  let manualCount = 0;
  const brkParts = list.map((b) => {
    const bAttrs: Record<string, string | number | boolean | undefined> = { id: b.id };
    if (b.min !== undefined) bAttrs.min = b.min;
    if (b.max !== undefined) bAttrs.max = b.max;
    if (b.manual) {
      bAttrs.man = 1;
      manualCount++;
    }
    if (b.pivot) bAttrs.pt = 1;
    return `<brk${attrs(bAttrs)}/>`;
  });
  return `<${tag} count="${list.length}" manualBreakCount="${manualCount}">${brkParts.join("")}</${tag}>`;
}

/**
 * Build the complete worksheet XML string.
 *
 * Zero-allocation fast path: directly concatenates XML string,
 * bypassing the intermediate object tree entirely.
 */
export function stringifyWorksheet(
  opts: WorksheetOptions,
  ctx: WorksheetContext,
  part?: { rootElement?: "worksheet" | "macrosheet" },
): string {
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
    part?.rootElement === "macrosheet"
      ? '<xm:macrosheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
        ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"' +
        ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"' +
        ' xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"' +
        ' xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"' +
        (opts.uid ? ` ${opts.uidPrefix ?? "xr"}:uid="${escapeXml(opts.uid)}"` : "") +
        ">"
      : '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
        ' mc:Ignorable="x14ac xr xr2 xr3"' +
        ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"' +
        ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"' +
        ' xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"' +
        ' xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"' +
        (opts.uid ? ` ${opts.uidPrefix ?? "xr"}:uid="${escapeXml(opts.uid)}"` : "") +
        ">",
  ];

  // Sheet properties (tabColor, outlinePr go here)
  const hasTabColor = !!opts.tabColor;
  const sp = opts.properties;
  const hasOutline =
    sp?.outlinePrPresent === true ||
    sp?.outlineSummaryBelow !== undefined ||
    sp?.outlineSummaryRight !== undefined ||
    sp?.outlineApplyStyles !== undefined ||
    sp?.outlineShowSymbols !== undefined;
  const hasSheetPrAttrs =
    sp &&
    (sp.codeName ||
      sp.syncHorizontal ||
      sp.syncVertical ||
      sp.syncRef ||
      sp.transitionEvaluation ||
      sp.transitionEntry ||
      sp.published !== undefined ||
      sp.filterMode !== undefined ||
      sp.enableFormatConditionsCalculation !== undefined);
  const hasPageSetUpPr =
    opts.pageSetup?.pageSetUpPrPresent === true ||
    !!opts.pageSetup?.fitToWidth ||
    !!opts.pageSetup?.fitToHeight ||
    opts.pageSetup?.fitToPage !== undefined ||
    opts.pageSetup?.autoPageBreaks !== undefined;
  if (sp !== undefined || hasTabColor || hasOutline || hasSheetPrAttrs || hasPageSetUpPr) {
    const prParts: string[] = [];
    const prAttrs: Record<string, string | number | boolean | undefined> = {};
    if (sp?.codeName) prAttrs.codeName = sp.codeName;
    if (sp?.syncHorizontal) prAttrs.syncHorizontal = 1;
    if (sp?.syncVertical) prAttrs.syncVertical = 1;
    if (sp?.syncRef) prAttrs.syncRef = sp.syncRef;
    if (sp?.transitionEvaluation) prAttrs.transitionEvaluation = 1;
    if (sp?.transitionEntry) prAttrs.transitionEntry = 1;
    // XSD defaults true — emit only the explicit-false form (0).
    if (sp?.published === false) prAttrs.published = 0;
    if (sp?.filterMode !== undefined) prAttrs.filterMode = sp.filterMode ? 1 : 0;
    if (sp?.enableFormatConditionsCalculation === false)
      prAttrs.enableFormatConditionsCalculation = 0;
    if (opts.tabColor) {
      const tc = typeof opts.tabColor === "string" ? { rgb: opts.tabColor } : opts.tabColor;
      const tcAttrs = colorAttributes(tc);
      prParts.push(`<tabColor${tcAttrs ? ` ${tcAttrs}` : ""}/>`);
    }
    if (hasOutline) {
      const outAttrs: Record<string, string | number | boolean | undefined> = {
        summaryBelow: 1,
        summaryRight: 1,
      };
      if (sp?.outlineSummaryBelow === false) outAttrs.summaryBelow = 0;
      if (sp?.outlineSummaryRight === false) outAttrs.summaryRight = 0;
      if (sp?.outlineApplyStyles) outAttrs.applyStyles = 1;
      if (sp?.outlineShowSymbols === false) outAttrs.showOutlineSymbols = 0;
      prParts.push(`<outlinePr${attrs(outAttrs)}/>`);
    }
    // pageSetUpPr (inside sheetPr when any of its attributes is requested).
    // fitToPage alone is meaningful: files parsed with fitToWidth/fitToHeight
    // at their XSD defaults (1/1, attributes omitted) must re-emit the flag.
    if (hasPageSetUpPr) {
      const psupAttrs: Record<string, string | number | boolean | undefined> = {};
      if (opts.pageSetup?.fitToPage !== undefined)
        psupAttrs.fitToPage = opts.pageSetup.fitToPage ? 1 : 0;
      else if (opts.pageSetup?.fitToWidth || opts.pageSetup?.fitToHeight) psupAttrs.fitToPage = 1;
      // autoPageBreaks defaults true — emit as written, explicit 0 included.
      if (opts.pageSetup?.autoPageBreaks !== undefined)
        psupAttrs.autoPageBreaks = opts.pageSetup.autoPageBreaks ? 1 : 0;
      prParts.push(`<pageSetUpPr${attrs(psupAttrs)}/>`);
    }
    const prAttrStr = Object.keys(prAttrs).length > 0 ? attrs(prAttrs) : "";
    p.push(
      prParts.length > 0
        ? `<sheetPr${prAttrStr}>${prParts.join("")}</sheetPr>`
        : selfCloseElement("sheetPr", prAttrStr),
    );
  }

  // Dimension — defines the used range of the sheet
  const maxRow = rows.length;
  let maxCol = 0;
  for (const row of rows) {
    if (row.cells && row.cells.length > maxCol) maxCol = row.cells.length;
  }
  if (opts.dimension) {
    p.push(`<dimension ref="${opts.dimension}"/>`);
  } else if (opts.dimensionPresent !== false && maxRow > 0 && maxCol > 0) {
    const dimRef = `A1:${defaultCellRef(maxRow, maxCol)}`;
    p.push(`<dimension ref="${dimRef}"/>`);
  }

  // Sheet views
  if (opts.freezePanes) {
    const fp = opts.freezePanes;
    const ySplit = fp.row ? fp.row : 0;
    const xSplit = fp.col ? fp.col : 0;
    const topRow = fp.row ? fp.row + 1 : 1;
    const leftCol = fp.col ? fp.col + 1 : 1;
    const topLeftCell = fp.topLeftCell ?? defaultCellRef(topRow, leftCol);
    const activePane =
      fp.activePane ??
      (ySplit > 0 && xSplit > 0 ? "bottomRight" : ySplit > 0 ? "bottomLeft" : "topRight");
    const state = fp.state ?? (fp.split ? "split" : "frozen");
    const paneAttrs =
      (fp.row !== undefined ? ` ySplit="${fp.row}"` : "") +
      (fp.col !== undefined ? ` xSplit="${fp.col}"` : "") +
      ` topLeftCell="${topLeftCell}" activePane="${activePane}" state="${state}"`;
    const svAttrs = buildSheetViewAttrs(opts.sheetView);
    const selections = (opts.selection ?? []).map(buildSelectionXml).join("");
    p.push(
      `<sheetViews><sheetView${svAttrs}>`,
      `<pane${paneAttrs}/>`,
      selections,
      opts.pivotSelection ? buildPivotSelectionXml(opts.pivotSelection) : "",
      "</sheetView></sheetViews>",
    );
  } else if (opts.sheetView !== undefined || opts.sheetViewsPresent !== false) {
    const svAttrs = buildSheetViewAttrs(opts.sheetView);
    const selections = (opts.selection ?? []).map(buildSelectionXml).join("");
    const innerXml =
      selections + (opts.pivotSelection ? buildPivotSelectionXml(opts.pivotSelection) : "");
    if (innerXml) {
      p.push(`<sheetViews><sheetView${svAttrs}>${innerXml}</sheetView></sheetViews>`);
    } else {
      p.push(`<sheetViews><sheetView${svAttrs}/></sheetViews>`);
    }
  }

  // Sheet format — default row height
  if (opts.sheetFormat === undefined) {
    p.push('<sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>');
  } else if (opts.sheetFormat) {
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
    if (sfp.dyDescent !== undefined) sfpAttrs["x14ac:dyDescent"] = sfp.dyDescent;
    p.push(`<sheetFormatPr${attrs(sfpAttrs)}/>`);
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
        colAttrs.width = col.widthRaw ?? col.width;
        if (col.customWidthDeclared !== false && col.customWidth === undefined)
          colAttrs.customWidth = 1;
      }
      // A column can carry customWidth="1" without a width (parse fills it);
      // preserve the explicit flag so round-trip does not drop the attribute.
      if (col.customWidth !== undefined) colAttrs.customWidth = col.customWidth ? 1 : 0;
      if (col.hidden !== undefined) colAttrs.hidden = col.hidden ? 1 : 0;
      if (col.outlineLevel !== undefined) {
        colAttrs.outlineLevel = col.outlineLevel;
      }
      if (col.collapsed !== undefined) colAttrs.collapsed = col.collapsed ? 1 : 0;
      if (col.bestFit) {
        colAttrs.bestFit = 1;
      }
      if (col.phonetic) {
        colAttrs.phonetic = 1;
      }
      if (col.style !== undefined) {
        colAttrs.style = col.style;
      }
      p.push(selfCloseElement("col", attrs(colAttrs)));
    }
    p.push("</cols>");
  }

  // Sheet data (rows + cells) — the hot path, shared with the streaming writer.
  p.push("<sheetData>");
  appendSheetDataRows(rows, p, sharedStrings, styles);
  p.push("</sheetData>");

  // Sheet calc properties (after sheetData per XSD sequence)
  if (opts.calculation) {
    const scAttrs: string[] = [];
    if (opts.calculation.fullCalcOnLoad) scAttrs.push('fullCalcOnLoad="1"');
    p.push(`<sheetCalcPr${scAttrs.length ? " " + scAttrs.join(" ") : ""}/>`);
  }

  // Sheet protection (after sheetData, before protectedRanges per XSD sequence)
  if (opts.protection) {
    p.push(stringifySheetProtectionXml(opts.protection));
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
      if (pr.securityDescriptor) {
        if (pr.securityDescriptorElement) {
          prParts.push(
            `<protectedRange${attrs(prAttrs)}><securityDescriptor>${escapeXml(pr.securityDescriptor)}</securityDescriptor></protectedRange>`,
          );
          continue;
        }
        prAttrs.securityDescriptor = pr.securityDescriptor;
        prParts.push(`<protectedRange${attrs(prAttrs)}/>`);
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
    if (opts.scenarios.sqref !== undefined) scAttrs.sqref = opts.scenarios.sqref;
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
          r: cell.reference,
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

  // Auto filter (CT_AutoFilter is shared with table — logic in auto-filter.ts)
  if (opts.autoFilter) {
    p.push(stringifyAutoFilter(opts.autoFilter));
  }

  // Sheet-level sort state (after autoFilter per XSD sequence)
  if (opts.sortState) {
    p.push(stringifySortStateXml(opts.sortState));
  }

  // Data consolidation
  if (opts.dataConsolidate) {
    const dc = opts.dataConsolidate;
    const dcAttrs: Record<string, string | number | boolean | undefined> = {};
    if (dc.function && dc.function !== "sum")
      dcAttrs.function = xsdConsolidateFunction.to(dc.function);
    if (dc.topLabels) dcAttrs.topLabels = 1;
    if (dc.leftLabels) dcAttrs.leftLabels = 1;
    if (dc.startLabels) dcAttrs.startLabels = 1;
    if (dc.link) dcAttrs.link = 1;
    const refsAttrs: Record<string, string | number> = {};
    if (dc.count !== undefined) refsAttrs.count = dc.count;
    const refsInner =
      dc.refs
        ?.map((input) => {
          const ref = typeof input === "string" ? { ref: input } : input;
          return selfCloseElement(
            "dataRef",
            attrs({
              ...(ref.ref !== undefined ? { ref: ref.ref } : {}),
              ...(ref.sheet !== undefined ? { sheet: ref.sheet } : {}),
              ...(ref.name !== undefined ? { name: ref.name } : {}),
              ...(ref.rId !== undefined ? { "r:id": ref.rId } : {}),
            }),
          );
        })
        .join("") ?? "";
    const refsXml = refsInner ? `<dataRefs${attrs(refsAttrs)}>${refsInner}</dataRefs>` : "";
    p.push(
      refsXml || Object.keys(dcAttrs).length > 0
        ? `<dataConsolidate${attrs(dcAttrs)}>${refsXml}</dataConsolidate>`
        : "<dataConsolidate/>",
    );
  }

  // Custom sheet views (after dataConsolidate per XSD sequence)
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
      if (csv.showRuler === false) csvAttrs.showRuler = 0;
      if (csv.topLeftCell !== undefined) csvAttrs.topLeftCell = csv.topLeftCell;
      if (csv.colorId !== undefined) csvAttrs.colorId = csv.colorId;
      let paneXml = "";
      if (csv.pane) {
        const pane = csv.pane;
        const paneAttrs: string[] = [];
        if (pane.row !== undefined) paneAttrs.push(`ySplit="${pane.row}"`);
        if (pane.col !== undefined) paneAttrs.push(`xSplit="${pane.col}"`);
        if (pane.topLeftCell !== undefined)
          paneAttrs.push(`topLeftCell="${escapeXml(pane.topLeftCell)}"`);
        if (pane.activePane !== undefined) paneAttrs.push(`activePane="${pane.activePane}"`);
        if (pane.state !== undefined || pane.split)
          paneAttrs.push(`state="${pane.state ?? (pane.split ? "split" : "frozen")}"`);
        paneXml = `<pane${paneAttrs.length > 0 ? ` ${paneAttrs.join(" ")}` : ""}/>`;
      }
      const selectionXml = (csv.selection ?? []).map(buildSelectionXml).join("");
      const pm = csv.pageMargins;
      const pageMarginsXml = pm
        ? `<pageMargins${attrs({
            left: convertToInch(pm.left ?? 0.75),
            right: convertToInch(pm.right ?? 0.75),
            top: convertToInch(pm.top ?? 1),
            bottom: convertToInch(pm.bottom ?? 1),
            header: convertToInch(pm.header ?? 0.5),
            footer: convertToInch(pm.footer ?? 0.5),
          })}/>`
        : "";
      const printOptionsXml = csv.printOptions ? stringifyPrintOptionsXml(csv.printOptions) : "";
      const pageSetupXml = csv.pageSetup ? stringifyPageSetupXml(csv.pageSetup) : "";
      const headerFooterXml = csv.headerFooter
        ? (stringifyHeaderFooterXml(csv.headerFooter) ?? "<headerFooter/>")
        : "";
      const autoFilterXml = stringifyAutoFilter(csv.autoFilter);
      const rowBreaksXml = csv.rowBreaks ? stringifyPageBreaksXml("rowBreaks", csv.rowBreaks) : "";
      const colBreaksXml = csv.colBreaks ? stringifyPageBreaksXml("colBreaks", csv.colBreaks) : "";
      const csvInner =
        paneXml +
        selectionXml +
        rowBreaksXml +
        colBreaksXml +
        pageMarginsXml +
        printOptionsXml +
        pageSetupXml +
        headerFooterXml +
        autoFilterXml;
      p.push(
        csvInner
          ? `<customSheetView${attrs(csvAttrs)}>${csvInner}</customSheetView>`
          : `<customSheetView${attrs(csvAttrs)}/>`,
      );
    }
    p.push("</customSheetViews>");
  }

  // Merge cells
  if (mergeCells.length > 0) {
    p.push(`<mergeCells count="${mergeCells.length}">`);
    for (const mc of mergeCells) {
      p.push(selfCloseElement("mergeCell", attrs({ ref: mc.ref })));
    }
    p.push("</mergeCells>");
  }

  // Phonetic properties (after mergeCells per XSD sequence)
  if (opts.phonetic) {
    const pp = opts.phonetic;
    const ppAttrs: Record<string, string | number> = { fontId: pp.fontId };
    if (pp.type && pp.type !== "fullwidthKatakana") ppAttrs.type = pp.type;
    if (pp.alignment && pp.alignment !== "left") ppAttrs.alignment = pp.alignment;
    p.push(selfCloseElement("phoneticPr", attrs(ppAttrs)));
  }

  // Conditional formatting
  const conditionalFormats = opts.conditionalFormats ?? [];
  if (conditionalFormats.length > 0) {
    for (const cf of conditionalFormats) {
      const cfAttrs = cf.pivot ? ` sqref="${cf.sqref}" pivot="1"` : ` sqref="${cf.sqref}"`;
      p.push(`<conditionalFormatting${cfAttrs}>`);
      for (const [ri, rule] of cf.rules.entries()) {
        const ruleAttrs: Record<string, string | number | boolean | undefined> = {
          type: rule.type,
          priority: rule.priority ?? ri + 1,
        };
        if (rule.operator) ruleAttrs.operator = rule.operator;
        if (rule.dxfId !== undefined) ruleAttrs.dxfId = rule.dxfId;
        if (rule.stopIfTrue) ruleAttrs.stopIfTrue = 1;
        if (rule.timePeriod) ruleAttrs.timePeriod = rule.timePeriod;
        if (rule.rank !== undefined) ruleAttrs.rank = rule.rank;
        if (rule.bottomRaw !== undefined) ruleAttrs.bottom = rule.bottomRaw;
        else if (rule.bottom) ruleAttrs.bottom = 1;
        if (rule.percentRaw !== undefined) ruleAttrs.percent = rule.percentRaw;
        else if (rule.percent) ruleAttrs.percent = 1;
        if (rule.text !== undefined) ruleAttrs.text = rule.text;
        if (rule.equalAverage) ruleAttrs.equalAverage = 1;
        if (rule.aboveAverage === false) ruleAttrs.aboveAverage = 0;
        if (rule.stdDev !== undefined) ruleAttrs.stdDev = rule.stdDev;
        const ruleExt = rule.ext ? `<extLst>${rule.ext}</extLst>` : "";

        const formulaXml = rule.formulas
          ?.map(
            (formula) =>
              `<formula${/^\s|\s$/.test(formula) ? ' xml:space="preserve"' : ""}>${escapeXml(formula)}</formula>`,
          )
          .join("");

        // Color scale
        if (rule.type === "colorScale" && rule.colorScale) {
          const cs = rule.colorScale;
          const inner: string[] = [];
          for (const v of cs.cfvo) {
            inner.push(buildCfvoXml(v));
          }
          for (const c of cs.colors) {
            inner.push(buildCfColorXml(c));
          }
          p.push(
            `<cfRule${attrs(ruleAttrs)}>${formulaXml ?? ""}<colorScale>${inner.join("")}</colorScale>${ruleExt}</cfRule>`,
          );
        }
        // Data bar
        else if (rule.type === "dataBar" && rule.dataBar) {
          const db = rule.dataBar;
          const inner: string[] = [];
          for (const v of db.cfvo) {
            inner.push(buildCfvoXml(v));
          }
          inner.push(buildCfColorXml(db.color));
          const dbAttrs: Record<string, string | number | boolean | undefined> = {};
          if (db.minLength !== undefined && db.minLength !== 10) dbAttrs.minLength = db.minLength;
          if (db.maxLength !== undefined && db.maxLength !== 90) dbAttrs.maxLength = db.maxLength;
          if (db.showValueRaw !== undefined) dbAttrs.showValue = db.showValueRaw;
          else if (db.showValue === false) dbAttrs.showValue = 0;
          const attrStr = Object.keys(dbAttrs).length > 0 ? attrs(dbAttrs) : "";
          p.push(
            `<cfRule${attrs(ruleAttrs)}>${formulaXml ?? ""}<dataBar${attrStr}>${inner.join("")}</dataBar>${ruleExt}</cfRule>`,
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
          if (is.showValueRaw !== undefined) isAttrs.showValue = is.showValueRaw;
          else if (is.showValue === false) isAttrs.showValue = 0;
          if (is.percentRaw !== undefined) isAttrs.percent = is.percentRaw;
          else if (is.percent === false) isAttrs.percent = 0;
          if (is.reverse) isAttrs.reverse = 1;
          const attrStr = Object.keys(isAttrs).length > 0 ? attrs(isAttrs) : "";
          p.push(
            `<cfRule${attrs(ruleAttrs)}>${formulaXml ?? ""}<iconSet${attrStr}>${inner.join("")}</iconSet>${ruleExt}</cfRule>`,
          );
        }
        // Standard rules (cellIs, containsText, expression, top10, aboveAverage)
        else {
          if (formulaXml) {
            p.push(`<cfRule${attrs(ruleAttrs)}>${formulaXml}${ruleExt}</cfRule>`);
          } else {
            p.push(
              ruleExt
                ? `<cfRule${attrs(ruleAttrs)}>${ruleExt}</cfRule>`
                : selfCloseElement("cfRule", attrs(ruleAttrs)),
            );
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
      if (dv.uid) dvAttrs["xr:uid"] = dv.uid;
      if (dv.type) dvAttrs.type = dv.type;
      if (dv.operator) dvAttrs.operator = dv.operator;
      if (dv.allowBlank !== undefined) dvAttrs.allowBlank = dv.allowBlank ? 1 : 0;
      if (dv.showErrorMessage !== undefined) dvAttrs.showErrorMessage = dv.showErrorMessage ? 1 : 0;
      if (dv.showInputMessage !== undefined) {
        dvAttrs.showInputMessage = dv.showInputMessage ? 1 : 0;
      }
      if (dv.errorTitle !== undefined) dvAttrs.errorTitle = dv.errorTitle;
      if (dv.error !== undefined) dvAttrs.error = dv.error;
      if (dv.promptTitle !== undefined) dvAttrs.promptTitle = dv.promptTitle;
      if (dv.prompt !== undefined) dvAttrs.prompt = dv.prompt;
      if (dv.errorStyle) dvAttrs.errorStyle = dv.errorStyle;
      if (dv.imeMode) dvAttrs.imeMode = dv.imeMode;
      if (dv.showDropDown !== undefined) dvAttrs.showDropDown = dv.showDropDown ? 1 : 0;
      const inner: string[] = [];
      if (dv.explicitList !== undefined)
        inner.push(
          `<mc:AlternateContent xmlns:x12ac="http://schemas.microsoft.com/office/spreadsheetml/2011/1/ac" ` +
            `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
            `<mc:Choice Requires="x12ac"><x12ac:list>${escapeXml(dv.explicitList)}</x12ac:list></mc:Choice>` +
            `<mc:Fallback><formula1>${escapeXml(dv.formula1 ?? "")}</formula1></mc:Fallback>` +
            `</mc:AlternateContent>`,
        );
      else if (dv.formula1 !== undefined)
        inner.push(
          `<formula1${/^\s|\s$/.test(dv.formula1) ? ' xml:space="preserve"' : ""}>${escapeXml(dv.formula1)}</formula1>`,
        );
      if (dv.formula2 !== undefined)
        inner.push(
          `<formula2${/^\s|\s$/.test(dv.formula2) ? ' xml:space="preserve"' : ""}>${escapeXml(dv.formula2)}</formula2>`,
        );
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
      // CT_Hyperlink's @r:id and @location are independent optional attrs —
      // an external workbook plus an internal jump target is a legal pair.
      if (hl.url !== undefined) {
        hlIdx++;
        hlAttrs["r:id"] = hl.relationshipId ?? `rId${hlIdx}`;
      }
      if (hl.location !== undefined) {
        hlAttrs.location = hl.location;
      }
      if (hl.tooltip) hlAttrs.tooltip = hl.tooltip;
      if (hl.display) hlAttrs.display = hl.display;
      if (hl.uid) hlAttrs["xr:uid"] = hl.uid;
      p.push(selfCloseElement("hyperlink", attrs(hlAttrs)));
    }
    p.push("</hyperlinks>");
  }

  // Print options
  if (opts.printOptions) {
    p.push(stringifyPrintOptionsXml(opts.printOptions));
  }

  if (opts.pageMargins) {
    const pm = opts.pageMargins;
    const raw = (value: number | UniversalMeasure | undefined, rawValue?: string) =>
      rawValue ?? (value === undefined ? "" : convertToInch(value));
    p.push(
      `<pageMargins${attrs({
        ...(pm.left !== undefined || pm.leftRaw !== undefined
          ? { left: raw(pm.left, pm.leftRaw) }
          : {}),
        ...(pm.right !== undefined || pm.rightRaw !== undefined
          ? { right: raw(pm.right, pm.rightRaw) }
          : {}),
        ...(pm.top !== undefined || pm.topRaw !== undefined ? { top: raw(pm.top, pm.topRaw) } : {}),
        ...(pm.bottom !== undefined || pm.bottomRaw !== undefined
          ? { bottom: raw(pm.bottom, pm.bottomRaw) }
          : {}),
        ...(pm.header !== undefined || pm.headerRaw !== undefined
          ? { header: raw(pm.header, pm.headerRaw) }
          : { header: 0.5 }),
        ...(pm.footer !== undefined || pm.footerRaw !== undefined
          ? { footer: raw(pm.footer, pm.footerRaw) }
          : { footer: 0.5 }),
      })}/>`,
    );
  }

  // Page setup — fitToPage/autoPageBreaks are pageSetUpPr fields stashed on
  // the same object (emitted inside sheetPr above); alone they must not
  // conjure a <pageSetup> the source never had.
  if (
    opts.pageSetup &&
    Object.keys(opts.pageSetup).some(
      (k) =>
        k !== "fitToPage" &&
        k !== "autoPageBreaks" &&
        k !== "pageSetUpPrPresent" &&
        k !== "pageSetupPresent",
    )
  ) {
    p.push(stringifyPageSetupXml(opts.pageSetup));
  }

  // Header/footer
  if (opts.headerFooter) {
    const hfXml = stringifyHeaderFooterXml(opts.headerFooter);
    p.push(hfXml ?? "<headerFooter/>");
  }

  // Row breaks (after headerFooter per XSD sequence), then column breaks
  if (rowBreaks.length > 0) p.push(stringifyPageBreaksXml("rowBreaks", rowBreaks));
  if (colBreaks.length > 0) p.push(stringifyPageBreaksXml("colBreaks", colBreaks));

  // Custom properties (CT_CustomProperties, after colBreaks per XSD sequence)
  if (customProperties.length > 0) {
    const cpParts: string[] = ["<customProperties>"];
    for (const cp of customProperties) {
      cpParts.push(`<customPr name="${escapeXml(cp.name)}" r:id="${escapeXml(cp.rId)}"/>`);
    }
    cpParts.push("</customProperties>");
    p.push(cpParts.join(""));
  }

  // Cell watches
  if (cellWatches.length > 0) {
    p.push("<cellWatches>");
    for (const cw of cellWatches) {
      p.push(`<cellWatch r="${escapeXml(cw.reference)}"/>`);
    }
    p.push("</cellWatches>");
  }

  // Ignored errors (after cellWatches per XSD sequence)
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

  // Cell smart tags (after ignoredErrors per XSD sequence)
  if (opts.smartTags && opts.smartTags.length > 0) {
    const stParts: string[] = ["<smartTags>"];
    for (const cst of opts.smartTags) {
      stParts.push(`<cellSmartTags r="${escapeXml(cst.reference)}">`);
      for (const st of cst.smartTags) {
        const stAttrs: string[] = [`type="${st.type}"`];
        if (st.deleted) stAttrs.push('deleted="1"');
        if (st.xmlBased) stAttrs.push('xmlBased="1"');
        const prXml = st.properties
          ? st.properties
              .map(
                (pr) => `<cellSmartTagPr key="${escapeXml(pr.key)}" val="${escapeXml(pr.val)}"/>`,
              )
              .join("")
          : "";
        stParts.push(
          prXml
            ? `<cellSmartTag ${stAttrs.join(" ")}>${prXml}</cellSmartTag>`
            : `<cellSmartTag ${stAttrs.join(" ")}/>`,
        );
      }
      stParts.push("</cellSmartTags>");
    }
    stParts.push("</smartTags>");
    p.push(stParts.join(""));
  }

  // Drawing / legacy drawing / table parts placeholders — the compiler owns
  // these parts' relationships; it replaces the markers it needs and strips
  // the rest, so the sheet-level order lives here in one place.
  p.push("<!--DRAWING-->");
  p.push("<!--LEGACY_DRAWING-->");

  // Legacy drawing in header/footer
  if (opts.legacyDrawingHF) {
    p.push(`<legacyDrawingHF r:id="${escapeXml(opts.legacyDrawingHF)}"/>`);
  }

  // Drawing in header/footer (after legacyDrawingHF per XSD sequence); attrs()
  // drops undefined so unset position fields stay absent.
  if (opts.drawingHF) {
    const dhf = opts.drawingHF;
    const dhfAttrs: Record<string, string | number | boolean | undefined> = {
      "r:id": dhf.rId,
      lho: dhf.lho,
      lhe: dhf.lhe,
      lhf: dhf.lhf,
      cho: dhf.cho,
      che: dhf.che,
      chf: dhf.chf,
      rho: dhf.rho,
      rhe: dhf.rhe,
      rhf: dhf.rhf,
      lfo: dhf.lfo,
      lfe: dhf.lfe,
      lff: dhf.lff,
      cfo: dhf.cfo,
      cfe: dhf.cfe,
      cff: dhf.cff,
      rfo: dhf.rfo,
      rfe: dhf.rfe,
      rff: dhf.rff,
    };
    p.push(selfCloseElement("drawingHF", attrs(dhfAttrs)));
  }

  // Background picture placeholder — compiler replaces with <picture r:id="rIdN"/>
  if (opts.backgroundImage) {
    p.push("<!--BACKGROUND_PICTURE-->");
  }

  /** CT_Marker corner inside an objectPr/controlPr anchor. */
  function anchorMarkerXml(tag: string, m: AnchorMarkerOptions): string {
    return (
      `<${tag}><xdr:col>${m.col}</xdr:col><xdr:colOff>${m.colOff ?? 0}</xdr:colOff>` +
      `<xdr:row>${m.row}</xdr:row><xdr:rowOff>${m.rowOff ?? 0}</xdr:rowOff></${tag}>`
    );
  }

  /** anchor element inside objectPr/controlPr (from/to corners, 0-based). */
  function embeddedAnchorXml(anchor: ObjectAnchorOptions): string {
    // The worksheet root does not declare xdr:, so the anchor declares it
    // locally — the marker children (xdr:col…) stay well-formed.
    const XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    const attrs = [
      `xmlns:xdr="${XDR_NS}"`,
      ...(anchor.moveWithCells ? ['moveWithCells="1"'] : []),
      ...(anchor.sizeWithCells ? ['sizeWithCells="1"'] : []),
    ].join(" ");
    return `<anchor ${attrs}>${anchorMarkerXml("from", anchor.from)}${anchorMarkerXml("to", anchor.to)}</anchor>`;
  }

  /**
   * mc:AlternateContent wrapper (Excel 2010+ form for worksheet OLE objects and
   * ActiveX controls): the Choice carries the full element, the Fallback the
   * bare one for older readers. Namespace declarations stay local so the
   * wrapper works regardless of what the worksheet root declares.
   */
  function wrapAlternateContent(
    choiceXml: string,
    fallbackXml: string,
    wrapped: boolean | undefined,
  ): string {
    if (!wrapped) return choiceXml;
    const AC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    const X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    return (
      `<mc:AlternateContent xmlns:mc="${AC_NS}">` +
      `<mc:Choice xmlns:x14="${X14_NS}" Requires="x14">${choiceXml}</mc:Choice>` +
      `<mc:Fallback>${fallbackXml}</mc:Fallback>` +
      `</mc:AlternateContent>`
    );
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
      let innerXml: string;
      if (ole.properties) {
        const opr = ole.properties;
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
        if (opr.iconRid) oprAttrs.push(`r:id="${escapeXml(opr.iconRid)}"`);
        const anchorXml = opr.anchor ? embeddedAnchorXml(opr.anchor) : "";
        innerXml = `<objectPr${oprAttrs.length ? " " + oprAttrs.join(" ") : ""}>${anchorXml}</objectPr>`;
      } else {
        innerXml = "";
      }
      const full = innerXml
        ? `<oleObject ${oleAttrs.join(" ")}>${innerXml}</oleObject>`
        : `<oleObject ${oleAttrs.join(" ")}/>`;
      oleParts.push(
        wrapAlternateContent(full, `<oleObject ${oleAttrs.join(" ")}/>`, ole.alternateContent),
      );
    }
    oleParts.push("</oleObjects>");
    p.push(oleParts.join(""));
  }

  // Controls (CT_Controls, after oleObjects per XSD sequence)
  if (controls.length > 0) {
    const controlXml: string[] = [];
    for (const control of controls) {
      if (control.rId === undefined) {
        throw new Error("xl/worksheets/controls/control: r:id: control relationship is not wired");
      }
      const controlAttrs: string[] = [
        `shapeId="${control.shapeId}"`,
        `r:id="${escapeXml(control.rId)}"`,
      ];
      if (control.name) controlAttrs.push(`name="${escapeXml(control.name)}"`);
      const attrsText = controlAttrs.join(" ");
      const properties = control.properties;
      let full: string;
      if (properties) {
        const controlPrAttrs: string[] = [];
        if (properties.locked === false) controlPrAttrs.push('locked="0"');
        if (properties.defaultSize === false) controlPrAttrs.push('defaultSize="0"');
        if (properties.print === false) controlPrAttrs.push('print="0"');
        if (properties.disabled) controlPrAttrs.push('disabled="1"');
        if (properties.recalcAlways) controlPrAttrs.push('recalcAlways="1"');
        if (properties.uiObject) controlPrAttrs.push('uiObject="1"');
        if (properties.autoFill === false) controlPrAttrs.push('autoFill="0"');
        if (properties.autoLine === false) controlPrAttrs.push('autoLine="0"');
        if (properties.autoPict === false) controlPrAttrs.push('autoPict="0"');
        if (properties.macro) controlPrAttrs.push(`macro="${escapeXml(properties.macro)}"`);
        if (properties.altText) controlPrAttrs.push(`altText="${escapeXml(properties.altText)}"`);
        if (properties.linkedCell)
          controlPrAttrs.push(`linkedCell="${escapeXml(properties.linkedCell)}"`);
        if (properties.listFillRange)
          controlPrAttrs.push(`listFillRange="${escapeXml(properties.listFillRange)}"`);
        if (properties.cf) controlPrAttrs.push(`cf="${escapeXml(properties.cf)}"`);
        if (properties.iconRid) controlPrAttrs.push(`r:id="${escapeXml(properties.iconRid)}"`);
        full = `<control ${attrsText}><controlPr ${controlPrAttrs.join(" ")}>${embeddedAnchorXml(
          properties.anchor,
        )}</controlPr></control>`;
      } else {
        full = `<control ${attrsText}/>`;
      }
      controlXml.push(
        wrapAlternateContent(full, `<control ${attrsText}/>`, control.alternateContent),
      );
    }
    const controlsXml = `<controls>${controlXml.join("")}</controls>`;
    p.push(
      opts.controlsAlternateContent
        ? wrapAlternateContent(controlsXml, controlsXml, true)
        : controlsXml,
    );
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

  p.push("<!--TABLE_PARTS-->");

  // Extension list (extLst, last per XSD sequence)
  if (opts.ext) {
    p.push(`<extLst>${opts.ext}</extLst>`);
  }

  p.push(part?.rootElement === "macrosheet" ? "</xm:macrosheet>" : "</worksheet>");
  return p.join("");
}

// ── Stringify helpers ──

function buildCfvoXml(cfvo: CfvoOptions): string {
  const a: Record<string, string | number | boolean | undefined> = { type: cfvo.type };
  if (cfvo.val !== undefined) a.val = cfvo.val;
  if (cfvo.gteRaw !== undefined) a.gte = cfvo.gteRaw;
  else if (cfvo.gte === false) a.gte = 0;
  return `<cfvo${attrs(a)}/>`;
}

/** Serialize a CT_Color. */
function buildCfColorXml(color: ColorOptions): string {
  const a = colorAttributes(color);
  return `<color${a ? ` ${a}` : ""}/>`;
}

export function buildSheetViewAttrs(sv?: SheetViewOptions): string {
  const svMap: Record<string, string | number | boolean | undefined> = {
    workbookViewId: 0,
  };
  if (sv?.tabSelected !== undefined) svMap.tabSelected = sv.tabSelected ? 1 : 0;
  // Omit tabSelected otherwise: only the active sheet carries it (Excel uses
  // workbookView activeTab), so injecting it on every sheet marks all active.
  if (sv?.showGridLines !== undefined) svMap.showGridLines = sv.showGridLines ? 1 : 0;
  if (sv?.showRowColHeaders !== undefined) svMap.showRowColHeaders = sv.showRowColHeaders ? 1 : 0;
  if (sv?.showZeros !== undefined) svMap.showZeros = sv.showZeros ? 1 : 0;
  if (sv?.zoomScale !== undefined) svMap.zoomScale = sv.zoomScale;
  if (sv?.rightToLeft !== undefined) svMap.rightToLeft = sv.rightToLeft ? 1 : 0;
  if (sv?.windowProtection !== undefined) svMap.windowProtection = sv.windowProtection ? 1 : 0;
  if (sv?.showFormulas !== undefined) svMap.showFormulas = sv.showFormulas ? 1 : 0;
  if (sv?.showRuler !== undefined) svMap.showRuler = sv.showRuler ? 1 : 0;
  if (sv?.showOutlineSymbols !== undefined)
    svMap.showOutlineSymbols = sv.showOutlineSymbols ? 1 : 0;
  if (sv?.defaultGridColor !== undefined) svMap.defaultGridColor = sv.defaultGridColor ? 1 : 0;
  if (sv?.showWhiteSpace !== undefined) svMap.showWhiteSpace = sv.showWhiteSpace ? 1 : 0;
  if (sv?.view) svMap.view = sv.view;
  if (sv?.topLeftCell) svMap.topLeftCell = sv.topLeftCell;
  if (sv?.colorId !== undefined) svMap.colorId = sv.colorId;
  if (sv?.zoomScaleNormal !== undefined) svMap.zoomScaleNormal = sv.zoomScaleNormal;
  if (sv?.zoomScaleSheetLayoutView !== undefined)
    svMap.zoomScaleSheetLayoutView = sv.zoomScaleSheetLayoutView;
  if (sv?.zoomScalePageLayoutView !== undefined)
    svMap.zoomScalePageLayoutView = sv.zoomScalePageLayoutView;
  return attrs(svMap);
}

function buildPivotSelectionXml(ps: PivotSelectionOptions): string {
  const psAttrs: string[] = [];
  if (ps.pane) psAttrs.push(`pane="${ps.pane}"`);
  if (ps.showHeader) psAttrs.push('showHeader="1"');
  if (ps.label) psAttrs.push('label="1"');
  if (ps.data) psAttrs.push('data="1"');
  if (ps.extendable) psAttrs.push('extendable="1"');
  if (ps.count !== undefined) psAttrs.push(`count="${ps.count}"`);
  if (ps.axis) psAttrs.push(`axis="${ps.axis}"`);
  if (ps.dimension !== undefined) psAttrs.push(`dimension="${ps.dimension}"`);
  if (ps.start !== undefined) psAttrs.push(`start="${ps.start}"`);
  if (ps.min !== undefined) psAttrs.push(`min="${ps.min}"`);
  if (ps.max !== undefined) psAttrs.push(`max="${ps.max}"`);
  if (ps.activeRow !== undefined) psAttrs.push(`activeRow="${ps.activeRow}"`);
  if (ps.activeCol !== undefined) psAttrs.push(`activeCol="${ps.activeCol}"`);
  if (ps.previousRow !== undefined) psAttrs.push(`previousRow="${ps.previousRow}"`);
  if (ps.previousCol !== undefined) psAttrs.push(`previousCol="${ps.previousCol}"`);
  if (ps.click !== undefined) psAttrs.push(`click="${ps.click}"`);
  if (ps.rId) psAttrs.push(`r:id="${escapeXml(ps.rId)}"`);
  const areaXml = ps.pivotArea ? buildPivotAreaXml(ps.pivotArea) : "";
  const attrStr = psAttrs.join(" ");
  return areaXml
    ? `<pivotSelection ${attrStr}>${areaXml}</pivotSelection>`
    : `<pivotSelection ${attrStr}/>`;
}

function buildSelectionXml(sel: SelectionOptions): string {
  const selAttrs: Record<string, string | number | boolean | undefined> = {};
  if (sel.pane) selAttrs.pane = sel.pane;
  if (sel.activeCell) selAttrs.activeCell = sel.activeCell;
  if (sel.activeCellId !== undefined) selAttrs.activeCellId = sel.activeCellId;
  if (sel.sqref) selAttrs.sqref = sel.sqref;
  return `<selection${attrs(selAttrs)}/>`;
}

function buildFormulaString(cellFormula: string | FormulaOptions): string {
  const fOpts = typeof cellFormula === "string" ? { formula: cellFormula } : cellFormula;
  const fAttrs: Record<string, string | number | boolean | undefined> = {};
  if (fOpts.type && fOpts.type !== FormulaType.NORMAL) fAttrs.t = fOpts.type;
  if (fOpts.reference) fAttrs.ref = fOpts.reference;
  if (fOpts.sharedIndex !== undefined) fAttrs.si = fOpts.sharedIndex;
  if (fOpts.aca) fAttrs.aca = 1;
  if (fOpts.dt2D) fAttrs.dt2D = 1;
  if (fOpts.dtr) fAttrs.dtr = 1;
  if (fOpts.del1) fAttrs.del1 = 1;
  if (fOpts.del2) fAttrs.del2 = 1;
  if (fOpts.inputCell1) fAttrs.r1 = fOpts.inputCell1;
  if (fOpts.inputCell2) fAttrs.r2 = fOpts.inputCell2;
  if (fOpts.calculateCell) fAttrs.ca = 1;
  if (fOpts.arrayContext) fAttrs.bx = 1;
  if (/^\s|\s$/.test(fOpts.formula)) fAttrs["xml:space"] = "preserve";

  const hasContent = fOpts.formula !== undefined && fOpts.formula !== "";

  if (hasContent) {
    return `<f${attrs(fAttrs)}>${escapeXml(fOpts.formula)}</f>`;
  }
  if (Object.keys(fAttrs).length > 0) {
    return selfCloseElement("f", attrs(fAttrs));
  }
  return "";
}

/**
 * Serialize the sheetData rows into `out` (strings appended in order, caller
 * wraps with `<sheetData>`/`</sheetData>`). Shared by the full stringify and
 * the streaming writer so the row/cell serialization cannot drift — passing
 * `sharedStrings: undefined` switches string cells to `t="inlineStr"`, the
 * constant-memory mode used when streaming. `startRowNumber` seeds the
 * implicit row numbering for callers streaming a slice of the full row list.
 */
export function appendSheetDataRows(
  rows: RowOptions[],
  out: string[],
  sharedStrings: SharedStrings | undefined,
  styles: Styles | undefined,
  startRowNumber = 1,
): void {
  // Indexed loops (not `entries()` destructuring) — at 100k rows × 20 cells the
  // per-iteration [index, value] entry arrays were a dominant Scavenge source.
  for (let i = 0; i < rows.length; i++) {
    const rowOpts = rows[i]!;
    const rowNumber = rowOpts.rowNumber ?? startRowNumber + i;
    // Flat attribute assembly (no per-row Record + for-in) keeps young-gen
    // pressure near zero at 100k+ rows.
    // An undefined rowNumber preserves a source row without `@r` — the row
    // position stays implicit; fresh rows emit the computed number only when
    // the caller set it.
    let rowAttr = rowOpts.rowNumber === undefined ? "" : ` r="${rowNumber}"`;
    if (rowOpts.dyDescent !== undefined) {
      rowAttr += ` x14ac:dyDescent="${rowOpts.dyDescent}"`;
    }
    if (rowOpts.height !== undefined) {
      rowAttr += ` ht="${convertToPt(rowOpts.height)}"`;
      if (rowOpts.customHeight !== undefined)
        rowAttr += ` customHeight="${rowOpts.customHeight ? 1 : 0}"`;
    } else if (rowOpts.customHeight !== undefined) {
      rowAttr += ` customHeight="${rowOpts.customHeight ? 1 : 0}"`;
    }
    if (rowOpts.hidden !== undefined) rowAttr += ` hidden="${rowOpts.hidden ? 1 : 0}"`;
    if (rowOpts.spans) rowAttr += ` spans="${rowOpts.spans}"`;
    let hasStyle = false;
    if (typeof rowOpts.style === "number") {
      // Round-trip fallback: emit the carried cellXfs index verbatim.
      rowAttr += ` s="${rowOpts.style}"`;
      hasStyle = true;
    } else if (rowOpts.style !== undefined && styles) {
      rowAttr += ` s="${styles.register(rowOpts.style)}"`;
      hasStyle = true;
    }
    // A row style only takes effect when customFormat flags it (Excel always
    // pairs @s with customFormat="1").
    if (rowOpts.customFormat === undefined && hasStyle) rowAttr += ' customFormat="1"';
    else if (rowOpts.customFormat !== undefined)
      rowAttr += ` customFormat="${rowOpts.customFormat ? 1 : 0}"`;
    if (rowOpts.thickTop) rowAttr += ' thickTop="1"';
    if (rowOpts.thickBot) rowAttr += ' thickBot="1"';
    if (rowOpts.phonetic) rowAttr += ' ph="1"';
    if (rowOpts.outlineLevel !== undefined) rowAttr += ` outlineLevel="${rowOpts.outlineLevel}"`;
    if (rowOpts.collapsed !== undefined) rowAttr += ` collapsed="${rowOpts.collapsed ? 1 : 0}"`;

    const cells = rowOpts.cells;
    if (cells) {
      out.push(`<row${rowAttr}>`);
      for (let j = 0; j < cells.length; j++) {
        const cell = cells[j]!;
        const ref = cell.reference ?? defaultCellRef(rowNumber, j + 1);
        const cellStr = buildCellString(ref, cell, sharedStrings, styles);
        if (cellStr) out.push(cellStr);
      }
      out.push("</row>");
    } else {
      out.push(`<row${rowAttr}/>`);
    }
  }
}

/**
 * Replace the first occurrence of a placeholder marker inside the sheet tail
 * region (between the sheetData close tag and `</worksheet>`). stringify emits
 * `<sheetData></sheetData>` even when empty, so the anchor always hits near the
 * end of the string, and cell text is XML-escaped so marker bytes cannot occur
 * inside the sheet body — scanning the small tail avoids walking a multi-MB
 * sheet both for hits and for misses. A missing marker returns the input
 * unchanged with zero copying.
 */
export function editSheetTailMarker(xml: string, marker: string, replacement: string): string {
  const tailStart = findSheetTailStart(xml);
  if (tailStart === -1) return xml;
  const tail = xml.slice(tailStart);
  const at = tail.indexOf(marker);
  if (at === -1) return xml;
  return xml.slice(0, tailStart) + tail.slice(0, at) + replacement + tail.slice(at + marker.length);
}

function findSheetTailStart(xml: string): number {
  // Backward search hits immediately: the close tag is near the string end.
  // The self-closing form only guards a future empty-sheet emission change.
  let at = xml.lastIndexOf("</sheetData>");
  if (at === -1) at = xml.lastIndexOf("<sheetData/>");
  return at;
}

/**
 * Strip the placeholder comments stringifyWorksheet emits for parts whose
 * relationships the compiler owns (drawing, legacyDrawing, tableParts,
 * background picture). The compiler replaces the ones it needs and strips the
 * rest; callers that write worksheet XML without the compiler (patch
 * append/replace) always strip.
 */
export function stripWorksheetPlaceholders(xml: string): string {
  let out = editSheetTailMarker(xml, "<!--BACKGROUND_PICTURE-->", "");
  out = editSheetTailMarker(out, "<!--TABLE_PARTS-->", "");
  out = editSheetTailMarker(out, "<!--LEGACY_DRAWING-->", "");
  return editSheetTailMarker(out, "<!--DRAWING-->", "");
}

function buildCellString(
  ref: string,
  cell: CellOptions,
  sharedStrings?: SharedStrings,
  styles?: Styles,
): string {
  // Flat attribute assembly (r → s → t, matching the former Record insertion
  // order byte-for-byte). No per-cell Record + for-in enumeration — at 2M cells
  // the dynamic-shape objects were a dominant Scavenge source.
  // An undefined reference preserves a source cell without `@r` — position
  // stays implicit, same as rows.
  const rAttr = cell.reference !== undefined ? ` r="${ref}"` : "";
  let sAttr = "";
  if (cell.styleRaw !== undefined) {
    // Source lexical form (e.g. an empty `s=""`); round-trip only.
    sAttr = ` s="${escapeXml(cell.styleRaw)}"`;
  } else if (typeof cell.style === "number") {
    // Round-trip fallback: emit the carried cellXfs index verbatim.
    sAttr = ` s="${cell.style}"`;
  } else if (cell.style !== undefined && styles) {
    sAttr = ` s="${styles.register(cell.style)}"`;
  }
  let mdAttr = "";
  if (cell.cellMetadataId !== undefined) mdAttr += ` cm="${cell.cellMetadataId}"`;
  if (cell.valueMetadataId !== undefined) mdAttr += ` vm="${cell.valueMetadataId}"`;
  if (cell.phonetic) mdAttr += ' ph="1"';

  const value = cell.value;

  // Error value (t="e") — emitted verbatim, never routed into the SST.
  if (cell.error !== undefined) {
    const errorV = cell.formula
      ? `${buildFormulaString(cell.formula)}<v>${escapeXml(cell.error)}</v>`
      : `<v>${escapeXml(cell.error)}</v>`;
    return `<c${rAttr}${sAttr}${mdAttr} t="e">${errorV}</c>`;
  }

  const rawValue =
    typeof value === "number" || typeof value === "string" ? cell.valueRaw : undefined;
  const rawPreserve = rawValue !== undefined && /^\s|\s$/.test(rawValue);
  const valueElement = (content: string, preserve = false): string =>
    `<v${preserve ? ' xml:space="preserve"' : ""}>${content}</v>`;

  // Formula path — formula takes precedence; value is the cached result.
  if (cell.formula) {
    const fStr = buildFormulaString(cell.formula);
    const fTypeAttr = cell.typeRaw !== undefined ? ` t="${escapeXml(cell.typeRaw)}"` : "";
    if (value === null || value === undefined) {
      return `<c${rAttr}${sAttr}${mdAttr}${fTypeAttr}>${fStr}</c>`;
    }
    let vStr = "";
    let tAttr = "";
    if (typeof value === "number") {
      if (cell.typeRaw === "n") tAttr = ' t="n"';
      vStr = valueElement(rawValue !== undefined ? escapeXml(rawValue) : `${value}`, rawPreserve);
    } else if (typeof value === "boolean") {
      tAttr = ' t="b"';
      vStr = `<v>${value ? 1 : 0}</v>`;
    } else if (typeof value === "string") {
      tAttr = cell.typeRaw === "d" ? ' t="d"' : ' t="str"';
      vStr = valueElement(escapeXml(rawValue ?? value), rawValue !== undefined);
    } else if (value instanceof Date) {
      vStr = `<v>${dateToSerialNumber(value)}</v>`;
    }
    if (vStr) {
      return `<c${rAttr}${sAttr}${mdAttr}${tAttr}>${fStr}${vStr}</c>`;
    }
    return `<c${rAttr}${sAttr}${mdAttr}>${fStr}</c>`;
  }

  if (value === null || value === undefined) {
    const tAttr = cell.typeRaw !== undefined ? ` t="${escapeXml(cell.typeRaw)}"` : "";
    if (cell.reference !== undefined || cell.style !== undefined || mdAttr || tAttr) {
      return `<c${rAttr}${sAttr}${mdAttr}${tAttr}/>`;
    }
    return "";
  }

  // Rich text value (RichTextOptions)
  if (typeof value === "object" && !(value instanceof Date)) {
    if (cell.inline) {
      const inlineV =
        cell.inlineValueRaw !== undefined ? `<v>${escapeXml(cell.inlineValueRaw)}</v>` : "";
      return `<c${rAttr}${sAttr}${mdAttr} t="inlineStr">${inlineV}<is>${buildRstXml(value)}</is></c>`;
    }
    if (sharedStrings) {
      const idx = sharedStrings.registerRich(value);
      return `<c${rAttr}${sAttr}${mdAttr} t="s"><v>${idx}</v></c>`;
    }
    return `<c${rAttr}${sAttr}${mdAttr} t="inlineStr"><is>${buildRstXml(value)}</is></c>`;
  }

  // Date cell (CT_Cell `t="d"`) — the value stays its source ISO lexical form.
  if (cell.typeRaw === "d" && typeof value === "string") {
    return `<c${rAttr}${sAttr}${mdAttr} t="d">${valueElement(
      escapeXml(rawValue ?? value),
      rawValue !== undefined,
    )}</c>`;
  }

  if (typeof value === "string") {
    if (cell.typeRaw === "str") {
      return `<c${rAttr}${sAttr}${mdAttr} t="str">${valueElement(
        escapeXml(rawValue ?? value),
        rawValue !== undefined,
      )}</c>`;
    }
    if (cell.inline) {
      return `<c${rAttr}${sAttr}${mdAttr} t="inlineStr"><is>${tElement(value)}</is></c>`;
    }
    // Round-trip: keep the source SST reference so duplicate si entries stay
    // stable; fresh cells register normally.
    // A finite sharedIndex carries the source table reference verbatim —
    // including packages whose SST part is missing, where the literal index
    // is still the source-faithful form.
    if (cell.sharedIndex !== undefined) {
      return `<c${rAttr}${sAttr}${mdAttr} t="s"><v>${cell.sharedIndex}</v></c>`;
    }
    if (sharedStrings) {
      if (value === "") return `<c${rAttr}${sAttr}${mdAttr} t="s"><v/></c>`;
      const idx = sharedStrings.register(value);
      return `<c${rAttr}${sAttr}${mdAttr} t="s"><v>${idx}</v></c>`;
    }
    return `<c${rAttr}${sAttr}${mdAttr} t="inlineStr"><is>${tElement(value)}</is></c>`;
  }

  if (typeof value === "number") {
    return `<c${rAttr}${sAttr}${mdAttr}${cell.typeRaw === "n" ? ' t="n"' : ""}>${valueElement(
      rawValue !== undefined ? escapeXml(rawValue) : `${value}`,
      rawPreserve,
    )}</c>`;
  }

  if (typeof value === "boolean") {
    return `<c${rAttr}${sAttr}${mdAttr} t="b"><v>${value ? 1 : 0}</v></c>`;
  }

  if (value instanceof Date) {
    const serial = dateToSerialNumber(value);
    return `<c${rAttr}${sAttr}${mdAttr}><v>${serial}</v></c>`;
  }

  return "";
}

function defaultCellRef(row: number, col: number): string {
  return columnToLetter(col) + row;
}

// ── Shared page-setup helpers (worksheet / dialogsheet / chartsheet) ──

/** Stringify CT_PageSetup attributes (worksheet + dialogsheet pageSetup). */
export function stringifyPageSetupXml(ps: PageSetupOptions): string {
  const psAttrs: Record<string, string | number | boolean | undefined> = {};
  if (ps.paperSize !== undefined) psAttrs.paperSize = ps.paperSize;
  if (ps.orientation !== undefined) psAttrs.orientation = ps.orientation;
  if (ps.scale !== undefined) psAttrs.scale = ps.scale;
  if (ps.fitToWidth !== undefined) psAttrs.fitToWidth = ps.fitToWidth;
  if (ps.fitToHeight !== undefined) psAttrs.fitToHeight = ps.fitToHeight;
  if (ps.pageOrder !== undefined) psAttrs.pageOrder = ps.pageOrder;
  if (ps.useFirstPageNumber !== undefined)
    psAttrs.useFirstPageNumber = ps.useFirstPageNumber ? 1 : 0;
  if (ps.firstPageNumber !== undefined) psAttrs.firstPageNumber = ps.firstPageNumber;
  // ST_PositiveUniversalMeasure requires a unit suffix; a bare number means mm.
  if (ps.paperHeight !== undefined)
    psAttrs.paperHeight =
      typeof ps.paperHeight === "number" ? `${ps.paperHeight}mm` : ps.paperHeight;
  if (ps.paperWidth !== undefined)
    psAttrs.paperWidth = typeof ps.paperWidth === "number" ? `${ps.paperWidth}mm` : ps.paperWidth;
  // XSD default true — emit only the explicit-false form (0).
  if (ps.usePrinterDefaults !== undefined)
    psAttrs.usePrinterDefaults = ps.usePrinterDefaults ? 1 : 0;
  if (ps.blackAndWhite !== undefined) psAttrs.blackAndWhite = ps.blackAndWhite ? 1 : 0;
  if (ps.draft !== undefined) psAttrs.draft = ps.draft ? 1 : 0;
  if (ps.cellComments !== undefined) psAttrs.cellComments = ps.cellComments;
  if (ps.errors !== undefined) psAttrs.errors = ps.errors;
  if (ps.horizontalDpi !== undefined) psAttrs.horizontalDpi = ps.horizontalDpi;
  if (ps.verticalDpi !== undefined) psAttrs.verticalDpi = ps.verticalDpi;
  if (ps.copies !== undefined) psAttrs.copies = ps.copies;
  if (ps.printerSettingsRId) psAttrs["r:id"] = ps.printerSettingsRId;
  return selfCloseElement("pageSetup", attrs(psAttrs));
}

/** Stringify a CT_PrintOptions element (worksheet + dialogsheet). */
export function stringifyPrintOptionsXml(po: PrintOptions): string {
  const poAttrs: Record<string, string | number | boolean | undefined> = {};
  if (po.horizontalCentered !== undefined)
    poAttrs.horizontalCentered = po.horizontalCentered ? 1 : 0;
  if (po.verticalCentered !== undefined) poAttrs.verticalCentered = po.verticalCentered ? 1 : 0;
  if (po.headings !== undefined) poAttrs.headings = po.headings ? 1 : 0;
  if (po.gridLines !== undefined) poAttrs.gridLines = po.gridLines ? 1 : 0;
  if (po.gridLinesSet !== undefined) poAttrs.gridLinesSet = po.gridLinesSet ? 1 : 0;
  return selfCloseElement("printOptions", attrs(poAttrs));
}

/** Stringify a CT_HeaderFooter element; undefined when it carries no content. */
export function stringifyHeaderFooterXml(hf: HeaderFooterOptions): string | undefined {
  const hfAttrs: Record<string, string | number | boolean | undefined> = {};
  if (hf.differentOddEven !== undefined) hfAttrs.differentOddEven = hf.differentOddEven ? 1 : 0;
  if (hf.differentFirst !== undefined) hfAttrs.differentFirst = hf.differentFirst ? 1 : 0;
  if (hf.scaleWithDocRaw !== undefined) hfAttrs.scaleWithDoc = hf.scaleWithDocRaw;
  else if (hf.scaleWithDoc === false) hfAttrs.scaleWithDoc = 0;
  if (hf.alignWithMarginsRaw !== undefined) hfAttrs.alignWithMargins = hf.alignWithMarginsRaw;
  else if (hf.alignWithMargins === false) hfAttrs.alignWithMargins = 0;
  const inner: string[] = [];
  const headerFooterPart = (name: HeaderFooterPartName, value: string | undefined): void => {
    if (value === undefined) return;
    // Replay the source xml:space form verbatim (Excel omits the attribute
    // even for trailing spaces); fresh authoring falls back to whitespace
    // detection so preserved spacing still round-trips by meaning.
    const recordedSpace = hf.xmlSpaceByPart?.[name];
    const xmlSpace =
      recordedSpace === "absent"
        ? undefined
        : (recordedSpace ?? (/^\s|\s$/.test(value) ? "preserve" : undefined));
    const preserve = xmlSpace ? ` xml:space="${xmlSpace}"` : "";
    inner.push(`<${name}${preserve}>${escapeXml(value)}</${name}>`);
  };
  headerFooterPart("oddHeader", hf.oddHeader);
  headerFooterPart("oddFooter", hf.oddFooter);
  headerFooterPart("evenHeader", hf.evenHeader);
  headerFooterPart("evenFooter", hf.evenFooter);
  headerFooterPart("firstHeader", hf.firstHeader);
  headerFooterPart("firstFooter", hf.firstFooter);
  if (inner.length > 0) {
    return `<headerFooter${attrs(hfAttrs)}>${inner.join("")}</headerFooter>`;
  }
  // Any explicit attribute (including scaleWithDoc/alignWithMargins = false)
  // keeps the element — dropping it would silently revert those to defaults.
  if (Object.keys(hfAttrs).length > 0) {
    return selfCloseElement("headerFooter", attrs(hfAttrs));
  }
  return undefined;
}

/** Stringify a CT_SheetProtection element (worksheet + dialogsheet). */
export function stringifySheetProtectionXml(prot: SheetProtectionOptions): string {
  const protAttrs: Record<string, string | number | boolean | undefined> = {};
  if (prot.passwordHashRaw !== undefined) {
    // Source legacy hash carries through verbatim — re-hashing would corrupt it.
    protAttrs.password = prot.passwordHashRaw;
  } else if (prot.password) {
    protAttrs.password = hashPassword(prot.password);
  }
  // Auto-derive modern hash when password provided without explicit hashValue
  let derived: ReturnType<typeof derivePasswordHash> | undefined;
  if (
    prot.password !== undefined &&
    prot.passwordHashRaw === undefined &&
    prot.hashValue === undefined
  ) {
    derived = derivePasswordHash(prot.password);
  }
  protAttrs.algorithmName = prot.algorithmName ?? derived?.algorithmName;
  protAttrs.hashValue = prot.hashValue ?? derived?.hashValue;
  protAttrs.saltValue = prot.saltValue ?? derived?.saltValue;
  if (prot.spinCount !== undefined) protAttrs.spinCount = prot.spinCount;
  else if (derived) protAttrs.spinCount = derived.spinCount;
  if (prot.sheet) protAttrs.sheet = 1;
  if (prot.objects) protAttrs.objects = 1;
  if (prot.scenariosRaw !== undefined) protAttrs.scenarios = prot.scenariosRaw;
  else if (prot.scenarios) protAttrs.scenarios = 1;
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
  return selfCloseElement("sheetProtection", attrs(protAttrs));
}
