/**
 * Worksheet — stringify implementation for xl/worksheets/sheet{n}.xml.
 *
 * Zero-allocation fast path: directly concatenates XML string, bypassing
 * the intermediate object tree entirely.
 *
 * @module
 */
import { convertToInch, convertToPt, derivePasswordHash } from "@office-open/core";
import { attrs, escapeXml, selfCloseElement } from "@office-open/xml";
import { columnToLetter, dateToSerialNumber, hashPassword } from "@util/index";

import { stringifyAutoFilter } from "../auto-filter";
import { buildPivotAreaXml } from "../pivot-table/stringify";
import { buildRstXml } from "../shared-strings";
import type { SharedStrings } from "../shared-strings";
import type { Styles } from "../styles";
import { FormulaType } from "./types";
import type {
  CellOptions,
  CfvoOptions,
  FormulaOptions,
  PivotSelectionOptions,
  RowOptions,
  SelectionOptions,
  SheetViewOptions,
  WorksheetContext,
  WorksheetOptions,
} from "./types";

/**
 * Build the complete worksheet XML string.
 *
 * Zero-allocation fast path: directly concatenates XML string,
 * bypassing the intermediate object tree entirely.
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
  const hasPageSetUpPr =
    !!opts.pageSetup?.fitToWidth ||
    !!opts.pageSetup?.fitToHeight ||
    !!opts.pageSetup?.fitToPage ||
    !!opts.pageSetup?.autoPageBreaks;
  if (hasTabColor || hasOutline || hasSheetPrAttrs || hasPageSetUpPr) {
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
      if (opts.pageSetup?.fitToWidth || opts.pageSetup?.fitToHeight || opts.pageSetup?.fitToPage)
        psupAttrs.fitToPage = 1;
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
  if (opts.dimension) {
    p.push(`<dimension ref="${opts.dimension}"/>`);
  } else if (maxRow > 0 && maxCol > 0) {
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
    const topLeftCell = defaultCellRef(topRow, leftCol);
    const activePane =
      ySplit > 0 && xSplit > 0 ? "bottomRight" : ySplit > 0 ? "bottomLeft" : "topRight";
    const svAttrs = buildSheetViewAttrs(opts.sheetView);
    p.push(
      `<sheetViews><sheetView${svAttrs}>`,
      `<pane ySplit="${ySplit}" xSplit="${xSplit}" topLeftCell="${topLeftCell}" activePane="${activePane}" state="frozen"/>`,
      opts.selection ? buildSelectionXml(opts.selection) : "",
      opts.pivotSelection ? buildPivotSelectionXml(opts.pivotSelection) : "",
      "</sheetView></sheetViews>",
    );
  } else {
    const svAttrs = buildSheetViewAttrs(opts.sheetView);
    const innerXml =
      (opts.selection ? buildSelectionXml(opts.selection) : "") +
      (opts.pivotSelection ? buildPivotSelectionXml(opts.pivotSelection) : "");
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
    p.push('<sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>');
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
      // A column can carry customWidth="1" without a width (parse fills it);
      // preserve the explicit flag so round-trip does not drop the attribute.
      if (col.customWidth) colAttrs.customWidth = 1;
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

  // Sheet data (rows + cells) — the hot path, shared with the streaming writer.
  p.push("<sheetData>");
  appendSheetDataRows(rows, p, sharedStrings, styles);
  p.push("</sheetData>");

  // Sheet calc properties (after sheetData per XSD sequence)
  if (opts.sheetCalcPr) {
    const scAttrs: string[] = [];
    if (opts.sheetCalcPr.fullCalcOnLoad) scAttrs.push('fullCalcOnLoad="1"');
    p.push(`<sheetCalcPr${scAttrs.length ? " " + scAttrs.join(" ") : ""}/>`);
  }

  // Row breaks (after sheetCalcPr per XSD sequence)
  if (rowBreaks.length > 0) {
    let manualCount = 0;
    const brkParts = rowBreaks.map((b) => {
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
    p.push(
      `<rowBreaks count="${rowBreaks.length}" manualBreakCount="${manualCount}">${brkParts.join("")}</rowBreaks>`,
    );
  }

  // Column breaks
  if (colBreaks.length > 0) {
    let manualCount = 0;
    const brkParts = colBreaks.map((b) => {
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
    p.push(
      `<colBreaks count="${colBreaks.length}" manualBreakCount="${manualCount}">${brkParts.join("")}</colBreaks>`,
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
      p.push(`<cellWatch r="${escapeXml(cw.reference)}"/>`);
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
      if (pr.securityDescriptor) prAttrs.securityDescriptor = pr.securityDescriptor;
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

  // Merge cells
  if (mergeCells.length > 0) {
    p.push(`<mergeCells count="${mergeCells.length}">`);
    for (const mc of mergeCells) {
      p.push(selfCloseElement("mergeCell", attrs({ ref: mc.ref })));
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
        if (rule.bottom) ruleAttrs.bottom = 1;
        if (rule.percent) ruleAttrs.percent = 1;
        if (rule.text !== undefined) ruleAttrs.text = rule.text;
        if (rule.equalAverage) ruleAttrs.equalAverage = 1;
        if (rule.aboveAverage === false) ruleAttrs.aboveAverage = 0;
        if (rule.stdDev !== undefined) ruleAttrs.stdDev = rule.stdDev;

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
      // CT_Hyperlink's @r:id and @location are independent optional attrs —
      // an external workbook plus an internal jump target is a legal pair.
      if (hl.url !== undefined) {
        hlIdx++;
        hlAttrs["r:id"] = `rId${hlIdx}`;
      }
      if (hl.location !== undefined) {
        hlAttrs.location = hl.location;
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

  if (opts.pageMargins) {
    const pm = opts.pageMargins;
    p.push(
      `<pageMargins${attrs({
        left: convertToInch(pm.left ?? 0.75),
        right: convertToInch(pm.right ?? 0.75),
        top: convertToInch(pm.top ?? 1),
        bottom: convertToInch(pm.bottom ?? 1),
        header: convertToInch(pm.header ?? 0.5),
        footer: convertToInch(pm.footer ?? 0.5),
      })}/>`,
    );
  } else {
    p.push('<pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>');
  }

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
      if (c.formula) prAttrs.push(`cf="${escapeXml(c.formula)}"`);
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
  // Omit tabSelected otherwise: only the active sheet carries it (Excel uses
  // workbookView activeTab), so injecting it on every sheet marks all active.
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
    let rowAttr = ` r="${rowNumber}"`;
    if (rowOpts.height !== undefined) {
      rowAttr += ` ht="${convertToPt(rowOpts.height)}" customHeight="1"`;
    }
    if (rowOpts.hidden) rowAttr += ' hidden="1"';
    if (rowOpts.spans) rowAttr += ` spans="${rowOpts.spans}"`;
    if (rowOpts.customFormat) rowAttr += ' customFormat="1"';
    if (rowOpts.thickTop) rowAttr += ' thickTop="1"';
    if (rowOpts.thickBot) rowAttr += ' thickBot="1"';
    if (rowOpts.phonetic) rowAttr += ' ph="1"';

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

function buildCellString(
  ref: string,
  cell: CellOptions,
  sharedStrings?: SharedStrings,
  styles?: Styles,
): string {
  // Flat attribute assembly (r → s → t, matching the former Record insertion
  // order byte-for-byte). No per-cell Record + for-in enumeration — at 2M cells
  // the dynamic-shape objects were a dominant Scavenge source.
  const rAttr = ` r="${ref}"`;
  let sAttr = "";
  if (typeof cell.style === "number") {
    // Round-trip fallback: emit the carried cellXfs index verbatim.
    sAttr = ` s="${cell.style}"`;
  } else if (cell.style !== undefined && styles) {
    sAttr = ` s="${styles.register(cell.style)}"`;
  }
  let mdAttr = "";
  if (cell.cellMetadataId !== undefined) mdAttr += ` cm="${cell.cellMetadataId}"`;
  if (cell.valueMetadataId !== undefined) mdAttr += ` vm="${cell.valueMetadataId}"`;

  const value = cell.value;

  // Formula path — formula takes precedence; value is the cached result.
  if (cell.formula) {
    const fStr = buildFormulaString(cell.formula);
    if (value === null || value === undefined) {
      return `<c${rAttr}${sAttr}${mdAttr}>${fStr}</c>`;
    }
    let vStr = "";
    let tAttr = "";
    if (typeof value === "number") {
      vStr = `<v>${value}</v>`;
    } else if (typeof value === "boolean") {
      tAttr = ' t="b"';
      vStr = `<v>${value ? 1 : 0}</v>`;
    } else if (typeof value === "string") {
      tAttr = ' t="str"';
      vStr = `<v>${escapeXml(value)}</v>`;
    } else if (value instanceof Date) {
      vStr = `<v>${dateToSerialNumber(value)}</v>`;
    }
    if (vStr) {
      return `<c${rAttr}${sAttr}${mdAttr}${tAttr}>${fStr}${vStr}</c>`;
    }
    return `<c${rAttr}${sAttr}${mdAttr}>${fStr}</c>`;
  }

  if (value === null || value === undefined) {
    if (cell.style !== undefined) {
      return `<c${rAttr}${sAttr}${mdAttr}/>`;
    }
    return "";
  }

  // Rich text value (RichTextOptions)
  if (typeof value === "object" && !(value instanceof Date)) {
    if (sharedStrings) {
      const idx = sharedStrings.registerRich(value);
      return `<c${rAttr}${sAttr}${mdAttr} t="s"><v>${idx}</v></c>`;
    }
    return `<c${rAttr}${sAttr}${mdAttr} t="inlineStr"><is>${buildRstXml(value)}</is></c>`;
  }

  if (typeof value === "string") {
    if (sharedStrings) {
      const idx = sharedStrings.register(value);
      return `<c${rAttr}${sAttr}${mdAttr} t="s"><v>${idx}</v></c>`;
    }
    return `<c${rAttr}${sAttr}${mdAttr} t="inlineStr"><is><t>${escapeXml(value)}</t></is></c>`;
  }

  if (typeof value === "number") {
    return `<c${rAttr}${sAttr}${mdAttr}><v>${value}</v></c>`;
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
