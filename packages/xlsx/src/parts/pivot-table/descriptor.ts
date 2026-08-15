/**
 * PivotTable descriptor for XLSX — generates xl/pivotTables/pivotTable{N}.xml.
 *
 * Implements CT_pivotTableDefinition from sml.xsd.
 * Direct stringify/parse — no intermediate class.
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import type { CustomDescriptor, WriteContext } from "@office-open/core/descriptor";
import { attr, attrNum, findChild, textOf } from "@office-open/xml";

import type { PivotAreaOptions } from "../pivot/pivot-utils";
import { parsePivotArea } from "./parse";
import { stringifyPivotTable } from "./stringify";
import type {
  CalculatedItemParseResult,
  CalculatedMemberParseResult,
  ChartFormatParseResult,
  ConditionalFormatParseResult,
  DataFieldParseResult,
  HierarchyUsageParseResult,
  PageFieldParseResult,
  PivotFieldParseResult,
  PivotFilterParseResult,
  PivotFormatParseResult,
  PivotHierarchyParseResult,
  PivotTableDescriptorOptions,
  PivotTableParseResult,
} from "./types";

// ── Descriptor ──

export const pivotTableDesc: CustomDescriptor<
  PivotTableDescriptorOptions,
  WriteContext,
  PivotTableParseResult
> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return stringifyPivotTable(opts.options, opts.sourceData, opts.cacheId);
  },

  parse(el, _ctx) {
    const result: Partial<PivotTableParseResult> = {};

    // Root element attributes
    if (attr(el, "name")) result.name = attr(el, "name");
    if (attr(el, "cacheId") !== undefined) result.cacheId = attrNum(el, "cacheId") ?? 0;
    if (parseOnOff(attr(el, "dataOnRows"))) result.dataOnRows = true;
    if (String(attr(el, "showHeaders")) === "0") result.showHeaders = false;
    if (parseOnOff(attr(el, "showEmptyRow"))) result.showEmptyRow = true;
    if (parseOnOff(attr(el, "showEmptyCol"))) result.showEmptyCol = true;
    if (attr(el, "grandTotalCaption")) result.grandTotalCaption = attr(el, "grandTotalCaption");
    if (attr(el, "errorCaption")) result.errorCaption = attr(el, "errorCaption");
    if (parseOnOff(attr(el, "showError"))) result.showError = true;
    if (attr(el, "missingCaption")) result.missingCaption = attr(el, "missingCaption");
    if (String(attr(el, "showMissing")) === "0") result.showMissing = false;
    if (attr(el, "pageStyle")) result.pageStyle = attr(el, "pageStyle");
    if (attr(el, "pivotTableStyle")) result.pivotTableStyle = attr(el, "pivotTableStyle");
    if (attr(el, "tag")) result.tag = attr(el, "tag");
    if (String(attr(el, "showItems")) === "0") result.showItems = false;
    if (parseOnOff(attr(el, "editData"))) result.editData = true;
    if (parseOnOff(attr(el, "disableFieldList"))) result.disableFieldList = true;
    if (String(attr(el, "showCalcMbrs")) === "0") result.showCalcMbrs = false;
    if (parseOnOff(attr(el, "visualTotals"))) result.visualTotals = true;
    if (String(attr(el, "showMultipleLabel")) === "0") result.showMultipleLabel = false;
    if (String(attr(el, "showDataDropDown")) === "0") result.showDataDropDown = false;
    if (String(attr(el, "showDrill")) === "0") result.showDrill = false;
    if (parseOnOff(attr(el, "printDrill"))) result.printDrill = true;
    if (parseOnOff(attr(el, "showMemberPropertyTips"))) result.showMemberPropertyTips = true;
    if (String(attr(el, "showDataTips")) === "0") result.showDataTips = false;
    if (String(attr(el, "enableWizard")) === "0") result.enableWizard = false;
    if (String(attr(el, "enableDrill")) === "0") result.enableDrill = false;
    if (String(attr(el, "enableFieldProperties")) === "0") result.enableFieldProperties = false;
    const pageWrap = attrNum(el, "pageWrap");
    if (pageWrap !== undefined) result.pageWrap = pageWrap;
    if (parseOnOff(attr(el, "pageOverThenDown"))) result.pageOverThenDown = true;
    if (parseOnOff(attr(el, "subtotalHiddenItems"))) result.subtotalHiddenItems = true;
    if (parseOnOff(attr(el, "fieldPrintTitles"))) result.fieldPrintTitles = true;
    if (parseOnOff(attr(el, "mergeItem"))) result.mergeItem = true;
    if (String(attr(el, "showDropZones")) === "0") result.showDropZones = false;
    if (parseOnOff(attr(el, "published"))) result.published = true;
    if (String(attr(el, "gridDropZones")) === "0") result.gridDropZones = false;
    if (String(attr(el, "multipleFieldFilters")) === "0") result.multipleFieldFilters = false;
    if (attr(el, "rowHeaderCaption")) result.rowHeaderCaption = attr(el, "rowHeaderCaption");
    if (attr(el, "colHeaderCaption")) result.colHeaderCaption = attr(el, "colHeaderCaption");
    if (parseOnOff(attr(el, "fieldListSortAscending"))) result.fieldListSortAscending = true;
    if (parseOnOff(attr(el, "mdxSubqueries"))) result.mdxSubqueries = true;
    if (String(attr(el, "customListSort")) === "0") result.customListSort = false;
    if (parseOnOff(attr(el, "asteriskTotals"))) result.asteriskTotals = true;
    const dataPosition = attrNum(el, "dataPosition");
    if (dataPosition !== undefined) result.dataPosition = dataPosition;
    if (parseOnOff(attr(el, "immersive"))) result.immersive = true;
    if (attr(el, "vacatedStyle")) result.vacatedStyle = attr(el, "vacatedStyle");
    if (parseOnOff(attr(el, "chartFormat"))) result.chartFormat = true;
    if (String(attr(el, "preserveFormatting")) === "0") result.preserveFormatting = false;
    else if (parseOnOff(attr(el, "preserveFormatting"))) result.preserveFormatting = true;
    if (attr(el, "dataCaption")) result.dataCaption = attr(el, "dataCaption");

    // Location — store ref as string, plus extended counts
    const locEl = findChild(el, "location");
    if (locEl) {
      if (attr(locEl, "ref")) result.location = attr(locEl, "ref");
      const rpc = attrNum(locEl, "rowPageCount");
      if (rpc !== undefined) result.locationRowPageCount = rpc;
      const cpc = attrNum(locEl, "colPageCount");
      if (cpc !== undefined) result.locationColPageCount = cpc;
    }

    // PivotFields
    const pfEl = findChild(el, "pivotFields");
    if (pfEl) {
      const fields: PivotFieldParseResult[] = [];
      for (const fEl of pfEl.elements ?? []) {
        if (fEl.name !== "pivotField") continue;
        const field: PivotFieldParseResult = {};
        const axis = attr(fEl, "axis");
        if (axis) field.axis = axis;
        if (String(attr(fEl, "showAll")) === "0") field.showAll = false;
        else if (parseOnOff(attr(fEl, "showAll"))) field.showAll = true;
        if (parseOnOff(attr(fEl, "dataField"))) field.dataField = true;
        if (attr(fEl, "hierarchy")) field.hierarchy = attr(fEl, "hierarchy");
        if (String(attr(fEl, "dragToRow")) === "0") field.dragToRow = false;
        if (String(attr(fEl, "dragToCol")) === "0") field.dragToCol = false;
        if (String(attr(fEl, "dragToPage")) === "0") field.dragToPage = false;
        if (parseOnOff(attr(fEl, "dragToData"))) field.dragToData = true;
        if (String(attr(fEl, "dragOff")) === "0") field.dragOff = false;
        if (String(attr(fEl, "showDropDowns")) === "0") field.showDropDowns = false;
        if (parseOnOff(attr(fEl, "insertBlankRow"))) field.insertBlankRow = true;
        if (parseOnOff(attr(fEl, "showPropCell"))) field.showPropCell = true;
        if (parseOnOff(attr(fEl, "showPropTip"))) field.showPropTip = true;
        if (parseOnOff(attr(fEl, "showPropAsCaption"))) field.showPropAsCaption = true;
        if (String(attr(fEl, "compact")) === "0") field.compact = false;
        if (parseOnOff(attr(fEl, "outline"))) field.outline = true;
        if (String(attr(fEl, "subtotalTop")) === "0") field.subtotalTop = false;
        if (parseOnOff(attr(fEl, "includeNewItemsInFilter"))) field.includeNewItemsInFilter = true;
        if (parseOnOff(attr(fEl, "avgSubtotal"))) field.avgSubtotal = true;
        if (parseOnOff(attr(fEl, "countASubtotal"))) field.countASubtotal = true;
        if (parseOnOff(attr(fEl, "maxSubtotal"))) field.maxSubtotal = true;
        if (parseOnOff(attr(fEl, "minSubtotal"))) field.minSubtotal = true;
        if (parseOnOff(attr(fEl, "sumSubtotal"))) field.sumSubtotal = true;
        fields.push(field);
      }
      result.pivotFields = fields;
    }

    // DataFields
    const dfEl = findChild(el, "dataFields");
    if (dfEl) {
      const dataFields: DataFieldParseResult[] = [];
      for (const dEl of dfEl.elements ?? []) {
        if (dEl.name !== "dataField") continue;
        const df: DataFieldParseResult = {};
        if (attr(dEl, "name")) df.name = attr(dEl, "name");
        const fld = attrNum(dEl, "fld");
        if (fld !== undefined) df.fld = fld;
        if (attr(dEl, "subtotal")) df.subtotal = attr(dEl, "subtotal");
        if (attr(dEl, "showDataAs")) df.showDataAs = attr(dEl, "showDataAs");
        const baseField = attrNum(dEl, "baseField");
        if (baseField !== undefined) df.baseField = baseField;
        const baseItem = attrNum(dEl, "baseItem");
        if (baseItem !== undefined) df.baseItem = baseItem;
        if (attr(dEl, "numFmtId")) df.numFmtId = attr(dEl, "numFmtId");
        dataFields.push(df);
      }
      result.dataFields = dataFields;
    }

    // Row fields
    const rowFieldsEl = findChild(el, "rowFields");
    if (rowFieldsEl) {
      const rowFields: number[] = [];
      for (const f of rowFieldsEl.elements ?? []) {
        if (f.name === "field") {
          const x = attrNum(f, "x");
          if (x !== undefined) rowFields.push(x);
        }
      }
      result.rowFields = rowFields;
    }

    // Col fields
    const colFieldsEl = findChild(el, "colFields");
    if (colFieldsEl) {
      const colFields: number[] = [];
      for (const f of colFieldsEl.elements ?? []) {
        if (f.name === "field") {
          const x = attrNum(f, "x");
          if (x !== undefined) colFields.push(x);
        }
      }
      result.colFields = colFields;
    }

    // Page fields
    const pageFieldsEl = findChild(el, "pageFields");
    if (pageFieldsEl) {
      const pageFields: PageFieldParseResult[] = [];
      for (const pf of pageFieldsEl.elements ?? []) {
        if (pf.name !== "pageField") continue;
        const pfResult: PageFieldParseResult = {};
        const fld = attrNum(pf, "fld");
        if (fld !== undefined) pfResult.fld = fld;
        const hier = attrNum(pf, "hier");
        if (hier !== undefined) pfResult.hier = hier;
        if (attr(pf, "cap")) pfResult.cap = attr(pf, "cap");
        const item = attrNum(pf, "item");
        if (item !== undefined) pfResult.item = item;
        pageFields.push(pfResult);
      }
      result.pageFields = pageFields;
    }

    // Formats
    const formatsEl = findChild(el, "formats");
    if (formatsEl) {
      const formats: PivotFormatParseResult[] = [];
      for (const fmtEl of formatsEl.elements ?? []) {
        if (fmtEl.name !== "format") continue;
        const fmt: PivotFormatParseResult = {};
        if (attr(fmtEl, "action")) fmt.action = attr(fmtEl, "action");
        const dxfId = attrNum(fmtEl, "dxfId");
        if (dxfId !== undefined) fmt.dxfId = dxfId;
        const paEl = findChild(fmtEl, "pivotArea");
        if (paEl) fmt.pivotArea = parsePivotArea(paEl);
        formats.push(fmt);
      }
      result.formats = formats;
    }

    // ConditionalFormats (pivot-specific CT_ConditionalFormats)
    const condFormatsEl = findChild(el, "conditionalFormats");
    if (condFormatsEl) {
      const condFormats: ConditionalFormatParseResult[] = [];
      for (const cfEl of condFormatsEl.elements ?? []) {
        if (cfEl.name !== "conditionalFormat") continue;
        const cf: ConditionalFormatParseResult = {};
        if (attr(cfEl, "scope")) cf.scope = attr(cfEl, "scope");
        if (attr(cfEl, "type")) cf.type = attr(cfEl, "type");
        const priority = attrNum(cfEl, "priority");
        if (priority !== undefined) cf.priority = priority;
        const areasEl = findChild(cfEl, "pivotAreas");
        if (areasEl) {
          const areas: Partial<PivotAreaOptions>[] = [];
          for (const aEl of areasEl.elements ?? []) {
            if (aEl.name !== "pivotArea") continue;
            areas.push(parsePivotArea(aEl));
          }
          cf.pivotAreas = areas;
        }
        condFormats.push(cf);
      }
      result.conditionalFormats = condFormats;
    }

    // ChartFormats
    const chartFormatsEl = findChild(el, "chartFormats");
    if (chartFormatsEl) {
      const chartFormats: ChartFormatParseResult[] = [];
      for (const cfEl of chartFormatsEl.elements ?? []) {
        if (cfEl.name !== "chartFormat") continue;
        const cf: ChartFormatParseResult = {};
        const chart = attrNum(cfEl, "chart");
        if (chart !== undefined) cf.chart = chart;
        const format = attrNum(cfEl, "format");
        if (format !== undefined) cf.format = format;
        if (parseOnOff(attr(cfEl, "series"))) cf.series = true;
        const paEl = findChild(cfEl, "pivotArea");
        if (paEl) cf.pivotArea = parsePivotArea(paEl);
        chartFormats.push(cf);
      }
      result.chartFormats = chartFormats;
    }

    // PivotHierarchies
    const hierarchiesEl = findChild(el, "pivotHierarchies");
    if (hierarchiesEl) {
      const hierarchies: PivotHierarchyParseResult[] = [];
      for (const hEl of hierarchiesEl.elements ?? []) {
        if (hEl.name !== "pivotHierarchy") continue;
        const h: PivotHierarchyParseResult = {};
        if (parseOnOff(attr(hEl, "outline"))) h.outline = true;
        if (parseOnOff(attr(hEl, "multipleItemSelectionAllowed")))
          h.multipleItemSelectionAllowed = true;
        if (parseOnOff(attr(hEl, "subtotalTop"))) h.subtotalTop = true;
        if (String(attr(hEl, "showInFieldList")) === "0") h.showInFieldList = false;
        if (String(attr(hEl, "dragToRow")) === "0") h.dragToRow = false;
        if (String(attr(hEl, "dragToCol")) === "0") h.dragToCol = false;
        if (String(attr(hEl, "dragToPage")) === "0") h.dragToPage = false;
        if (parseOnOff(attr(hEl, "dragToData"))) h.dragToData = true;
        if (String(attr(hEl, "dragOff")) === "0") h.dragOff = false;
        if (parseOnOff(attr(hEl, "includeNewItemsInFilter"))) h.includeNewItemsInFilter = true;
        if (attr(hEl, "caption")) h.caption = attr(hEl, "caption");
        hierarchies.push(h);
      }
      result.pivotHierarchies = hierarchies;
    }

    // Filters
    const filtersEl = findChild(el, "filters");
    if (filtersEl) {
      const filters: PivotFilterParseResult[] = [];
      for (const fEl of filtersEl.elements ?? []) {
        if (fEl.name !== "filter") continue;
        const f: PivotFilterParseResult = {};
        const fld = attrNum(fEl, "fld");
        if (fld !== undefined) f.fld = fld;
        if (attr(fEl, "type")) f.type = attr(fEl, "type");
        const id = attrNum(fEl, "id");
        if (id !== undefined) f.id = id;
        const mpFld = attrNum(fEl, "mpFld");
        if (mpFld !== undefined) f.mpFld = mpFld;
        const evalOrder = attrNum(fEl, "evalOrder");
        if (evalOrder !== undefined) f.evalOrder = evalOrder;
        const imh = attrNum(fEl, "iMeasureHier");
        if (imh !== undefined) f.iMeasureHier = imh;
        const imf = attrNum(fEl, "iMeasureFld");
        if (imf !== undefined) f.iMeasureFld = imf;
        if (attr(fEl, "stringValue1")) f.stringValue1 = attr(fEl, "stringValue1");
        if (attr(fEl, "stringValue2")) f.stringValue2 = attr(fEl, "stringValue2");
        filters.push(f);
      }
      result.filters = filters;
    }

    // RowHierarchiesUsage
    const rhuEl = findChild(el, "rowHierarchiesUsage");
    if (rhuEl) {
      const usage: HierarchyUsageParseResult[] = [];
      for (const u of rhuEl.elements ?? []) {
        if (u.name === "rowHierarchyUsage") {
          usage.push({ hierarchyUsage: attrNum(u, "hierarchyUsage") ?? 0 });
        }
      }
      result.rowHierarchiesUsage = usage;
    }

    // ColHierarchiesUsage
    const chuEl = findChild(el, "colHierarchiesUsage");
    if (chuEl) {
      const usage: HierarchyUsageParseResult[] = [];
      for (const u of chuEl.elements ?? []) {
        if (u.name === "colHierarchyUsage") {
          usage.push({ hierarchyUsage: attrNum(u, "hierarchyUsage") ?? 0 });
        }
      }
      result.colHierarchiesUsage = usage;
    }

    // CalculatedItems
    const ciEl = findChild(el, "calculatedItems");
    if (ciEl) {
      const items: CalculatedItemParseResult[] = [];
      for (const iEl of ciEl.elements ?? []) {
        if (iEl.name !== "calculatedItem") continue;
        const item: CalculatedItemParseResult = {};
        const field = attrNum(iEl, "field");
        if (field !== undefined) item.field = field;
        const formulaEl = findChild(iEl, "formula");
        if (formulaEl) item.formula = textOf(formulaEl);
        const paEl = findChild(iEl, "pivotArea");
        if (paEl) item.pivotArea = parsePivotArea(paEl);
        items.push(item);
      }
      result.calculatedItems = items;
    }

    // CalculatedMembers
    const cmEl = findChild(el, "calculatedMembers");
    if (cmEl) {
      const members: CalculatedMemberParseResult[] = [];
      for (const mEl of cmEl.elements ?? []) {
        if (mEl.name !== "calculatedMember") continue;
        const m: CalculatedMemberParseResult = {};
        if (attr(mEl, "name")) m.name = attr(mEl, "name");
        const mdxEl = findChild(mEl, "mdx");
        if (mdxEl) m.mdx = textOf(mdxEl) ?? "";
        if (attr(mEl, "memberName")) m.memberName = attr(mEl, "memberName");
        if (attr(mEl, "hierarchy")) m.hierarchy = attr(mEl, "hierarchy");
        if (attr(mEl, "parent")) m.parent = attr(mEl, "parent");
        const solveOrder = attrNum(mEl, "solveOrder");
        if (solveOrder !== undefined) m.solveOrder = solveOrder;
        if (parseOnOff(attr(mEl, "set"))) m.set = true;
        members.push(m);
      }
      result.calculatedMembers = members;
    }

    // Style from pivotTableStyleInfo/@name (the standard location)
    const styleInfoEl = findChild(el, "pivotTableStyleInfo");
    if (styleInfoEl) {
      const styleName = attr(styleInfoEl, "name");
      if (styleName) result.style = styleName;
    } else if (attr(el, "styleName")) {
      result.style = attr(el, "styleName");
    }

    return result as PivotTableParseResult;
  },
};
