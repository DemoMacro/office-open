/**
 * PivotTable descriptor for XLSX — generates xl/pivotTables/pivotTable{N}.xml.
 *
 * Implements CT_pivotTableDefinition from sml.xsd.
 * Direct stringify/parse — no intermediate class.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attrs, escapeXml } from "@office-open/xml";
import { findChild, attr, attrNum } from "@office-open/xml";

import type {
  PivotTableOptions,
  PivotSourceData,
  PivotDataField,
  PivotHierarchyOptions,
  PivotAreaOptions,
  PivotFieldOverrideOptions,
} from "../../file/pivot/pivot-utils";
import { collectUniqueValues } from "../../file/pivot/pivot-utils";

// ── Types ──

export interface PivotTableDescriptorOptions {
  options: PivotTableOptions;
  sourceData: PivotSourceData;
  cacheId: number;
}

// ── Descriptor ──

export const pivotTableDesc: CustomDescriptor<PivotTableDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return stringifyPivotTable(opts.options, opts.sourceData, opts.cacheId);
  },

  parse(el, _ctx) {
    const result: Record<string, unknown> = {};

    if (attr(el, "name")) result.name = attr(el, "name");
    if (attr(el, "cacheId") !== undefined) result.cacheId = attrNum(el, "cacheId") ?? 0;

    const locEl = findChild(el, "location");
    if (locEl) result.location = attr(locEl, "ref") ?? "";

    const pfEl = findChild(el, "pivotFields");
    if (pfEl) {
      const fields: Record<string, unknown>[] = [];
      for (const fEl of pfEl.elements ?? []) {
        if (fEl.name !== "pivotField") continue;
        const field: Record<string, unknown> = {};
        const axis = attr(fEl, "axis");
        if (axis) field.axis = axis;
        fields.push(field);
      }
      result.pivotFields = fields;
    }

    // DataFields
    const dfEl = findChild(el, "dataFields");
    if (dfEl) {
      const dataFields: Record<string, unknown>[] = [];
      for (const dEl of dfEl.elements ?? []) {
        if (dEl.name !== "dataField") continue;
        const df: Record<string, unknown> = {};
        if (attr(dEl, "name")) df.name = attr(dEl, "name");
        const fld = attrNum(dEl, "fld");
        if (fld !== undefined) df.fld = fld;
        if (attr(dEl, "subtotal")) df.subtotal = attr(dEl, "subtotal");
        dataFields.push(df);
      }
      result.dataFields = dataFields;
    }

    if (attr(el, "styleName")) result.style = attr(el, "styleName");

    return result as Record<string, unknown>;
  },
};

// ── Stringify implementation ──

function stringifyPivotTable(o: PivotTableOptions, sd: PivotSourceData, cacheId: number): string {
  const fields = sd.fieldNames;
  const rowFieldNames = o.rows;
  const colFieldNames = o.columns ?? [];
  const dataFields = o.data;
  const style = o.style ?? "PivotStyleLight16";
  const location = o.location ?? "A3";
  const name = o.name ?? "PivotTable1";

  const rowFieldIndices = rowFieldNames.map((n) => fields.indexOf(n));
  const colFieldIndices = colFieldNames.map((n) => fields.indexOf(n));
  const dataFieldIndices = dataFields.map((df) => fields.indexOf(df.field));
  const pageFieldNames = o.pages ?? [];
  const pageFieldIndices = pageFieldNames.map((n) => fields.indexOf(n));

  const pivotFieldsXml = buildPivotFields(
    o,
    sd,
    rowFieldIndices,
    colFieldIndices,
    dataFieldIndices,
    pageFieldIndices,
  );
  const pageFieldsXml = buildPageFields(o, pageFieldIndices);
  const rowFieldsXml = buildRowFields(rowFieldIndices);
  const rowItemsXml = buildRowItems(sd, rowFieldIndices);
  const colFieldsXml = buildColFields(colFieldIndices);
  const colItemsXml = buildColItems(sd, colFieldIndices, dataFields);
  const dataFieldsXml = buildDataFields(dataFields, dataFieldIndices);

  const locationRef = computeLocationRef(
    sd,
    location,
    rowFieldIndices,
    colFieldIndices,
    dataFields,
  );

  const p: string[] = [];
  const defAttrs: string[] = [
    `name="${escapeXml(name)}"`,
    `cacheId="${cacheId}"`,
    'dataCaption="Values"',
    'updatedVersion="6"',
    'minRefreshableVersion="3"',
    'createdVersion="6"',
    'applyNumberFormats="0"',
    'applyBorderFormats="0"',
    'applyFontFormats="0"',
    'applyPatternFormats="0"',
    'applyAlignmentFormats="0"',
    'applyWidthHeightFormats="1"',
    'autoFormatId="0"',
    'useAutoFormatting="1"',
    'itemPrintTitles="1"',
    'indent="0"',
    'outline="1"',
    'outlineData="1"',
    'compact="1"',
    'compactData="1"',
    'rowGrandTotals="1"',
    'colGrandTotals="1"',
  ];
  if (o.dataOnRows) defAttrs.push('dataOnRows="1"');
  if (o.grandTotalCaption) defAttrs.push(`grandTotalCaption="${escapeXml(o.grandTotalCaption)}"`);
  if (o.errorCaption) defAttrs.push(`errorCaption="${escapeXml(o.errorCaption)}"`);
  if (o.showError) defAttrs.push('showError="1"');
  if (o.missingCaption) defAttrs.push(`missingCaption="${escapeXml(o.missingCaption)}"`);
  if (o.showMissing === false) defAttrs.push('showMissing="0"');
  if (o.pageStyle) defAttrs.push(`pageStyle="${escapeXml(o.pageStyle)}"`);
  if (o.pivotTableStyle) defAttrs.push(`pivotTableStyle="${escapeXml(o.pivotTableStyle)}"`);
  if (o.tag) defAttrs.push(`tag="${escapeXml(o.tag)}"`);
  if (o.showItems === false) defAttrs.push('showItems="0"');
  if (o.editData) defAttrs.push('editData="1"');
  if (o.disableFieldList) defAttrs.push('disableFieldList="1"');
  if (o.showCalcMbrs === false) defAttrs.push('showCalcMbrs="0"');
  if (o.visualTotals) defAttrs.push('visualTotals="1"');
  if (o.showMultipleLabel === false) defAttrs.push('showMultipleLabel="0"');
  if (o.showDataDropDown === false) defAttrs.push('showDataDropDown="0"');
  if (o.showDrill === false) defAttrs.push('showDrill="0"');
  if (o.printDrill) defAttrs.push('printDrill="1"');
  if (o.showMemberPropertyTips) defAttrs.push('showMemberPropertyTips="1"');
  if (o.showDataTips === false) defAttrs.push('showDataTips="0"');
  if (o.enableWizard === false) defAttrs.push('enableWizard="0"');
  if (o.enableDrill === false) defAttrs.push('enableDrill="0"');
  if (o.enableFieldProperties === false) defAttrs.push('enableFieldProperties="0"');
  if (o.pageWrap !== undefined) defAttrs.push(`pageWrap="${o.pageWrap}"`);
  if (o.pageOverThenDown) defAttrs.push('pageOverThenDown="1"');
  if (o.subtotalHiddenItems) defAttrs.push('subtotalHiddenItems="1"');
  if (o.fieldPrintTitles) defAttrs.push('fieldPrintTitles="1"');
  if (o.mergeItem) defAttrs.push('mergeItem="1"');
  if (o.showDropZones === false) defAttrs.push('showDropZones="0"');
  if (o.showEmptyRow) defAttrs.push('showEmptyRow="1"');
  if (o.showEmptyCol) defAttrs.push('showEmptyCol="1"');
  if (o.showHeaders === false) defAttrs.push('showHeaders="0"');
  if (o.published) defAttrs.push('published="1"');
  if (o.gridDropZones === false) defAttrs.push('gridDropZones="0"');
  if (o.multipleFieldFilters === false) defAttrs.push('multipleFieldFilters="0"');
  if (o.rowHeaderCaption) defAttrs.push(`rowHeaderCaption="${escapeXml(o.rowHeaderCaption)}"`);
  if (o.colHeaderCaption) defAttrs.push(`colHeaderCaption="${escapeXml(o.colHeaderCaption)}"`);
  if (o.fieldListSortAscending) defAttrs.push('fieldListSortAscending="1"');
  if (o.mdxSubqueries) defAttrs.push('mdxSubqueries="1"');
  if (o.customListSort === false) defAttrs.push('customListSort="0"');
  if (o.asteriskTotals) defAttrs.push('asteriskTotals="1"');
  if (o.dataPosition !== undefined) defAttrs.push(`dataPosition="${o.dataPosition}"`);
  if (o.immersive) defAttrs.push('immersive="1"');
  if (o.vacatedStyle) defAttrs.push(`vacatedStyle="${escapeXml(o.vacatedStyle)}"`);

  p.push(
    `<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ${defAttrs.join(" ")}>`,
  );

  // location
  const locAttrs: string[] = [
    `ref="${escapeXml(locationRef)}"`,
    `firstHeaderRow="1"`,
    `firstDataRow="${colFieldIndices.length + 1}"`,
    `firstDataCol="${rowFieldIndices.length}"`,
  ];
  if (o.locationColPageCount !== undefined)
    locAttrs.push(`colPageCount="${o.locationColPageCount}"`);
  if (o.locationRowPageCount !== undefined)
    locAttrs.push(`rowPageCount="${o.locationRowPageCount}"`);
  p.push(`<location ${locAttrs.join(" ")}/>`);

  p.push(pivotFieldsXml);
  p.push(rowFieldsXml);
  p.push(rowItemsXml);
  if (colFieldIndices.length > 0) p.push(colFieldsXml);
  p.push(colItemsXml);
  if (pageFieldIndices.length > 0) p.push(pageFieldsXml);
  if (dataFields.length > 0) p.push(dataFieldsXml);

  // formats
  if (o.formats && o.formats.length > 0) {
    const fmtParts: string[] = [`<formats count="${o.formats.length}">`];
    for (const fmt of o.formats) {
      const fmtAttrs: string[] = [];
      if (fmt.action && fmt.action !== "formatting") fmtAttrs.push(`action="${fmt.action}"`);
      if (fmt.dxfId !== undefined) fmtAttrs.push(`dxfId="${fmt.dxfId}"`);
      fmtParts.push(
        `<format${fmtAttrs.length ? " " + fmtAttrs.join(" ") : ""}>${buildPivotAreaXml(fmt.pivotArea)}</format>`,
      );
    }
    fmtParts.push("</formats>");
    p.push(fmtParts.join(""));
  }

  // chartFormats
  if (o.chartFormats && o.chartFormats.length > 0) {
    const cfParts: string[] = [`<chartFormats count="${o.chartFormats.length}">`];
    for (const cf of o.chartFormats) {
      const cfAttrs: string[] = [`chart="${cf.chart}"`, `format="${cf.format}"`];
      if (cf.series) cfAttrs.push('series="1"');
      const areaXml = cf.pivotArea ? buildPivotAreaXml(cf.pivotArea) : "";
      cfParts.push(`<chartFormat ${cfAttrs.join(" ")}>${areaXml}</chartFormat>`);
    }
    cfParts.push("</chartFormats>");
    p.push(cfParts.join(""));
  }

  // pivotHierarchies
  if (o.pivotHierarchies && o.pivotHierarchies.length > 0) {
    p.push(buildPivotHierarchies(o.pivotHierarchies));
  }

  // pivotTableStyleInfo
  p.push(
    `<pivotTableStyleInfo name="${escapeXml(style)}" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>`,
  );

  // filters
  if (o.filters && o.filters.length > 0) {
    const fParts: string[] = [`<filters count="${o.filters.length}">`];
    for (const f of o.filters) {
      const fAttrs: Record<string, string | number | boolean | undefined> = {
        fld: f.fld,
        type: f.type,
        id: f.id,
      };
      if (f.mpFld !== undefined) fAttrs.mpFld = f.mpFld;
      if (f.evalOrder !== undefined) fAttrs.evalOrder = f.evalOrder;
      fParts.push(`<filter${attrs(fAttrs)}><autoFilter></autoFilter></filter>`);
    }
    fParts.push("</filters>");
    p.push(fParts.join(""));
  }

  // rowHierarchiesUsage
  if (o.rowHierarchiesUsage && o.rowHierarchiesUsage.length > 0) {
    const rhu = o.rowHierarchiesUsage;
    p.push(
      `<rowHierarchiesUsage count="${rhu.length}">${rhu.map((h) => `<rowHierarchyUsage hierarchyUsage="${h.hierarchyUsage}"/>`).join("")}</rowHierarchiesUsage>`,
    );
  }

  // colHierarchiesUsage
  if (o.colHierarchiesUsage && o.colHierarchiesUsage.length > 0) {
    const chu = o.colHierarchiesUsage;
    p.push(
      `<colHierarchiesUsage count="${chu.length}">${chu.map((h) => `<colHierarchyUsage hierarchyUsage="${h.hierarchyUsage}"/>`).join("")}</colHierarchiesUsage>`,
    );
  }

  p.push("</pivotTableDefinition>");
  return p.join("");
}

// ── Stringify helpers ──

function buildFieldOverrideAttrs(fo: PivotFieldOverrideOptions): string {
  const a: string[] = [];
  if (fo.allDrilled) a.push('allDrilled="1"');
  if (fo.autoShow) a.push('autoShow="1"');
  if (fo.countSubtotal) a.push('countSubtotal="1"');
  if (fo.dataSourceSort) a.push('dataSourceSort="1"');
  if (fo.defaultAttributeDrillState) a.push('defaultAttributeDrillState="1"');
  if (fo.hiddenLevel) a.push('hiddenLevel="1"');
  if (fo.hideNewItems) a.push('hideNewItems="1"');
  if (fo.insertBlankRow) a.push('insertBlankRow="1"');
  if (fo.insertPageBreak) a.push('insertPageBreak="1"');
  if (fo.itemPageCount) a.push('itemPageCount="1"');
  if (fo.measureFilter) a.push('measureFilter="1"');
  if (fo.nonAutoSortDefault) a.push('nonAutoSortDefault="1"');
  if (fo.productSubtotal) a.push('productSubtotal="1"');
  if (fo.rankBy !== undefined) a.push(`rankBy="${fo.rankBy}"`);
  if (fo.serverField) a.push('serverField="1"');
  if (fo.showDropDowns) a.push('showDropDowns="1"');
  if (fo.showPropAsCaption) a.push('showPropAsCaption="1"');
  if (fo.showPropCell) a.push('showPropCell="1"');
  if (fo.showPropTip) a.push('showPropTip="1"');
  if (fo.stdDevPSubtotal) a.push('stdDevPSubtotal="1"');
  if (fo.stdDevSubtotal) a.push('stdDevSubtotal="1"');
  if (fo.subtotalCaption) a.push(`subtotalCaption="${escapeXml(fo.subtotalCaption)}"`);
  if (fo.topAutoShow) a.push('topAutoShow="1"');
  if (fo.uniqueMemberProperty) a.push('uniqueMemberProperty="1"');
  if (fo.varPSubtotal) a.push('varPSubtotal="1"');
  if (fo.varSubtotal) a.push('varSubtotal="1"');
  return a.join(" ");
}

function buildPivotFields(
  o: PivotTableOptions,
  sd: PivotSourceData,
  rowIndices: number[],
  colIndices: number[],
  dataIndices: number[],
  pageIndices: number[],
): string {
  const fieldNames = sd.fieldNames;
  const parts: string[] = [`<pivotFields count="${fieldNames.length}">`];

  for (let i = 0; i < fieldNames.length; i++) {
    const isRow = rowIndices.includes(i);
    const isCol = colIndices.includes(i);
    const isData = dataIndices.includes(i);
    const isPage = pageIndices.includes(i);
    const override = o.fieldOverrides?.find((fo) => fo.field === fieldNames[i]);
    const extraAttrs = override ? buildFieldOverrideAttrs(override) : "";

    if (isData) {
      const dataFieldIdx = dataIndices.indexOf(i);
      const df = o.data[dataFieldIdx];
      const dfAttrs: string[] = ['dataField="1"', 'showAll="0"'];
      if (extraAttrs) dfAttrs.push(extraAttrs);
      if (df?.showDataAs) dfAttrs.push(`showDataAs="${df.showDataAs}"`);
      if (df?.baseField !== undefined) dfAttrs.push(`baseField="${df.baseField}"`);
      if (df?.baseItem !== undefined) dfAttrs.push(`baseItem="${df.baseItem}"`);
      if (o.autoSortScope) {
        parts.push(
          `<pivotField ${dfAttrs.join(" ")}><autoSortScope>${buildPivotAreaXml(o.autoSortScope)}</autoSortScope></pivotField>`,
        );
      } else {
        parts.push(`<pivotField ${dfAttrs.join(" ")}/>`);
      }
    } else if (isRow) {
      const uniqueVals = collectUniqueValues(sd.records, i);
      const rAttrs = extraAttrs
        ? ` axis="axisRow" showAll="0" ${extraAttrs}`
        : ' axis="axisRow" showAll="0"';
      parts.push(`<pivotField${rAttrs}>`);
      parts.push(`<items count="${uniqueVals.length + 1}">`);
      for (let j = 0; j < uniqueVals.length; j++) parts.push(`<item x="${j}"/>`);
      parts.push(`<item t="default"${override?.defaultItemSd === false ? ' sd="0"' : ""}/>`);
      parts.push("</items></pivotField>");
    } else if (isCol) {
      const uniqueVals = collectUniqueValues(sd.records, i);
      const cAttrs = extraAttrs
        ? ` axis="axisCol" showAll="0" ${extraAttrs}`
        : ' axis="axisCol" showAll="0"';
      parts.push(`<pivotField${cAttrs}>`);
      parts.push(`<items count="${uniqueVals.length + 1}">`);
      for (let j = 0; j < uniqueVals.length; j++) parts.push(`<item x="${j}"/>`);
      parts.push(`<item t="default"${override?.defaultItemSd === false ? ' sd="0"' : ""}/>`);
      parts.push("</items></pivotField>");
    } else if (isPage) {
      const uniqueVals = collectUniqueValues(sd.records, i);
      const pAttrs = extraAttrs
        ? ` axis="axisPage" showAll="0" ${extraAttrs}`
        : ' axis="axisPage" showAll="0"';
      parts.push(`<pivotField${pAttrs}>`);
      parts.push(`<items count="${uniqueVals.length + 1}">`);
      for (let j = 0; j < uniqueVals.length; j++) parts.push(`<item x="${j}"/>`);
      parts.push(`<item t="default"${override?.defaultItemSd === false ? ' sd="0"' : ""}/>`);
      parts.push("</items></pivotField>");
    } else {
      const nAttrs = extraAttrs ? ` showAll="0" ${extraAttrs}` : ' showAll="0"';
      parts.push(`<pivotField${nAttrs}/>`);
    }
  }

  parts.push("</pivotFields>");
  return parts.join("");
}

function buildPageFields(o: PivotTableOptions, pageIndices: number[]): string {
  if (pageIndices.length === 0) return "";
  const parts: string[] = [`<pageFields count="${pageIndices.length}">`];
  for (let i = 0; i < pageIndices.length; i++) {
    const cap = o.pageCaptions?.[i];
    const capAttr = cap ? ` cap="${escapeXml(cap)}"` : "";
    parts.push(`<pageField fld="${pageIndices[i]}" hier="${i}"${capAttr}/>`);
  }
  parts.push("</pageFields>");
  return parts.join("");
}

function buildRowFields(rowIndices: number[]): string {
  if (rowIndices.length === 0) return '<rowFields count="0"/>';
  const parts: string[] = [`<rowFields count="${rowIndices.length}">`];
  for (const idx of rowIndices) parts.push(`<field x="${idx}"/>`);
  parts.push("</rowFields>");
  return parts.join("");
}

function buildRowItems(sd: PivotSourceData, rowIndices: number[]): string {
  if (rowIndices.length === 0) return '<rowItems count="1"><i/></rowItems>';

  const allUniqueCounts: number[] = [];
  for (const idx of rowIndices) {
    allUniqueCounts.push(collectUniqueValues(sd.records, idx).length);
  }

  if (rowIndices.length === 1) {
    const count = allUniqueCounts[0];
    const parts: string[] = [`<rowItems count="${count + 1}">`];
    for (let i = 0; i < count; i++) parts.push(`<i><x v="${i}"/></i>`);
    parts.push(`<i t="grand"><x/></i>`);
    parts.push("</rowItems>");
    return parts.join("");
  }

  const combos = cartesianOfCounts(allUniqueCounts);
  const rowItems: string[] = [];
  for (const combo of combos) {
    rowItems.push(`<i>${combo.map((v) => `<x v="${v}"/>`).join("")}</i>`);
  }
  rowItems.push(`<i t="grand">${rowIndices.map(() => "<x/>").join("")}</i>`);
  return `<rowItems count="${rowItems.length}">${rowItems.join("")}</rowItems>`;
}

function buildColFields(colIndices: number[]): string {
  if (colIndices.length === 0) return '<colFields count="0"/>';
  const parts: string[] = [`<colFields count="${colIndices.length}">`];
  for (const idx of colIndices) parts.push(`<field x="${idx}"/>`);
  parts.push("</colFields>");
  return parts.join("");
}

function buildColItems(
  sd: PivotSourceData,
  colIndices: number[],
  dataFields: PivotDataField[],
): string {
  if (colIndices.length > 0) {
    const allUniqueCounts: number[] = [];
    for (const idx of colIndices) {
      allUniqueCounts.push(collectUniqueValues(sd.records, idx).length);
    }
    const combos = cartesianOfCounts(allUniqueCounts);
    const items: string[] = [];
    for (const combo of combos) {
      items.push(`<i>${combo.map((v) => `<x v="${v}"/>`).join("")}</i>`);
    }
    items.push(`<i t="grand">${colIndices.map(() => "<x/>").join("")}</i>`);
    return `<colItems count="${items.length}">${items.join("")}</colItems>`;
  }
  if (dataFields.length > 1) {
    const items = dataFields.map((_, i) => `<i><x v="${i}"/></i>`);
    return `<colItems count="${items.length}">${items.join("")}</colItems>`;
  }
  return '<colItems count="1"><i/></colItems>';
}

function buildDataFields(dataFields: PivotDataField[], dataFieldIndices: number[]): string {
  if (dataFields.length === 0) return '<dataFields count="0"/>';
  const parts: string[] = [`<dataFields count="${dataFields.length}">`];
  for (let i = 0; i < dataFields.length; i++) {
    const df = dataFields[i];
    const subtotal = df.summarize ?? "sum";
    const name = df.name ?? `${subtotal === "sum" ? "Sum" : subtotal} of ${df.field}`;
    const dfAttrs: string[] = [
      `name="${escapeXml(name)}"`,
      `fld="${dataFieldIndices[i]}"`,
      `subtotal="${subtotal}"`,
    ];
    if (df.showDataAs) dfAttrs.push(`showDataAs="${df.showDataAs}"`);
    if (df.baseField !== undefined) dfAttrs.push(`baseField="${df.baseField}"`);
    if (df.baseItem !== undefined) dfAttrs.push(`baseItem="${df.baseItem}"`);
    parts.push(`<dataField ${dfAttrs.join(" ")}/>`);
  }
  parts.push("</dataFields>");
  return parts.join("");
}

function computeLocationRef(
  sd: PivotSourceData,
  location: string,
  rowFieldIndices: number[],
  colFieldIndices: number[],
  dataFields: PivotDataField[],
): string {
  const startCell = location.split(":")[0];
  const match = startCell.match(/^([A-Z]+)(\d+)$/);
  if (!match) return location;
  const startCol = match[1];
  const startRow = parseInt(match[2], 10);

  let rowCount = 1;
  if (rowFieldIndices.length > 0) {
    rowCount += collectUniqueValues(sd.records, rowFieldIndices[0]).length;
  }
  rowCount += 1;

  let colCount = Math.max(rowFieldIndices.length, 1);
  if (colFieldIndices.length > 0) {
    colCount += collectUniqueValues(sd.records, colFieldIndices[0]).length;
  } else if (dataFields.length > 1) {
    colCount += dataFields.length - 1;
  }
  colCount += 1;

  const endCol = colIndexToLetter(letterToColIndex(startCol) + colCount - 1);
  const endRow = startRow + rowCount - 1;
  return `${startCol}${startRow}:${endCol}${endRow}`;
}

function buildPivotHierarchies(hierarchies: PivotHierarchyOptions[]): string {
  const parts: string[] = [`<pivotHierarchies count="${hierarchies.length}">`];
  for (const h of hierarchies) {
    const hAttrs: string[] = [];
    if (h.outline) hAttrs.push('outline="1"');
    if (h.multipleItemSelectionAllowed) hAttrs.push('multipleItemSelectionAllowed="1"');
    if (h.subtotalTop) hAttrs.push('subtotalTop="1"');
    if (h.showInFieldList === false) hAttrs.push('showInFieldList="0"');
    if (h.dragToRow === false) hAttrs.push('dragToRow="0"');
    if (h.dragToCol === false) hAttrs.push('dragToCol="0"');
    if (h.dragToPage === false) hAttrs.push('dragToPage="0"');
    if (h.dragToData) hAttrs.push('dragToData="1"');
    if (h.dragOff === false) hAttrs.push('dragOff="0"');
    if (h.includeNewItemsInFilter) hAttrs.push('includeNewItemsInFilter="1"');
    if (h.caption) hAttrs.push(`caption="${escapeXml(h.caption)}"`);
    const inner =
      (h.memberProperties
        ? `<mps count="${h.memberProperties.length}">${h.memberProperties
            .map((mp) => {
              const mpAttrs: string[] = [`field="${mp.field}"`];
              if (mp.name !== undefined) mpAttrs.push(`name="${escapeXml(mp.name)}"`);
              if (mp.showCell) mpAttrs.push('showCell="1"');
              if (mp.showTip) mpAttrs.push('showTip="1"');
              if (mp.showAsCaption) mpAttrs.push('showAsCaption="1"');
              return `<mp ${mpAttrs.join(" ")}/>`;
            })
            .join("")}</mps>`
        : "") +
      (h.members
        ? `<members count="${h.members.length}">${h.members.map((m) => `<member name="${escapeXml(m.name)}"${m.level !== undefined ? ` level="${m.level}"` : ""}/>`).join("")}</members>`
        : "");
    if (inner) {
      parts.push(`<pivotHierarchy ${hAttrs.join(" ")}>${inner}</pivotHierarchy>`);
    } else {
      parts.push(`<pivotHierarchy ${hAttrs.join(" ")}/>`);
    }
  }
  parts.push("</pivotHierarchies>");
  return parts.join("");
}

function buildPivotAreaXml(area: PivotAreaOptions): string {
  const aAttrs: string[] = [];
  if (area.field !== undefined) aAttrs.push(`field="${area.field}"`);
  if (area.type) aAttrs.push(`type="${area.type}"`);
  if (area.dataOnly === false) aAttrs.push('dataOnly="0"');
  if (area.labelOnly) aAttrs.push('labelOnly="1"');
  if (area.grandRow) aAttrs.push('grandRow="1"');
  if (area.grandCol) aAttrs.push('grandCol="1"');
  if (area.cacheIndex) aAttrs.push('cacheIndex="1"');
  if (area.outline === false) aAttrs.push('outline="0"');
  if (area.offset) aAttrs.push(`offset="${escapeXml(area.offset)}"`);
  if (area.collapsedLevelsAreSubtotals) aAttrs.push('collapsedLevelsAreSubtotals="1"');
  if (area.axis) aAttrs.push(`axis="${area.axis}"`);
  if (area.fieldPosition !== undefined) aAttrs.push(`fieldPosition="${area.fieldPosition}"`);
  const refsXml = area.references ? buildPivotAreaReferences(area.references) : "";
  if (refsXml) return `<pivotArea ${aAttrs.join(" ")}>${refsXml}</pivotArea>`;
  return `<pivotArea ${aAttrs.join(" ")}/>`;
}

function buildPivotAreaReferences(
  refs: import("../../file/pivot/pivot-utils").PivotAreaReferenceOptions[],
): string {
  const parts: string[] = [`<references count="${refs.length}">`];
  for (const ref of refs) {
    const rAttrs: string[] = [];
    if (ref.field !== undefined) rAttrs.push(`field="${ref.field}"`);
    if (ref.count !== undefined) rAttrs.push(`count="${ref.count}"`);
    if (ref.selected === false) rAttrs.push('selected="0"');
    if (ref.byPosition) rAttrs.push('byPosition="1"');
    if (ref.relative) rAttrs.push('relative="1"');
    if (ref.defaultSubtotal) rAttrs.push('defaultSubtotal="1"');
    const xXml = ref.x ? ref.x.map((v) => `<x v="${v}"/>`).join("") : "";
    if (xXml) {
      parts.push(`<reference ${rAttrs.join(" ")}>${xXml}</reference>`);
    } else {
      parts.push(`<reference ${rAttrs.join(" ")}/>`);
    }
  }
  parts.push("</references>");
  return parts.join("");
}

function letterToColIndex(letters: string): number {
  let col = 0;
  for (let i = 0; i < letters.length; i++) col = col * 26 + (letters.charCodeAt(i) - 64);
  return col;
}

function colIndexToLetter(col: number): string {
  let result = "";
  let n = col;
  while (n > 0) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}

function cartesianOfCounts(counts: number[]): number[][] {
  if (counts.length === 0) return [[]];
  let result: number[][] = [[]];
  for (const count of counts) {
    const next: number[][] = [];
    for (const prefix of result) {
      for (let i = 0; i < count; i++) next.push([...prefix, i]);
    }
    result = next;
  }
  return result;
}
