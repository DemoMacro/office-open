/**
 * PivotTableDefinition XML generator.
 *
 * Generates xl/pivotTables/pivotTable{N}.xml.
 * Follows CT_pivotTableDefinition from sml.xsd.
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { attrs, escapeXml } from "@office-open/xml";

import type { PivotSourceData } from "./pivot-utils";
import type {
  PivotTableOptions,
  PivotDataField,
  PivotFilterOptions,
  CalculatedItemOptions,
  CalculatedMemberOptions,
  PivotHierarchyOptions,
  PivotConditionalFormatOptions,
  ChartFormatOptions,
  PivotAreaOptions,
  PivotFieldOverrideOptions,
} from "./pivot-utils";
import { collectUniqueValues } from "./pivot-utils";

export class PivotTableXml extends BaseXmlComponent {
  private readonly options: PivotTableOptions;
  private readonly sourceData: PivotSourceData;
  private readonly cacheId: number;

  public constructor(options: PivotTableOptions, sourceData: PivotSourceData, cacheId: number) {
    super("pivotTableDefinition");
    this.options = options;
    this.sourceData = sourceData;
    this.cacheId = cacheId;
  }

  public override toXml(_context: Context): string {
    const o = this.options;
    const sd = this.sourceData;
    const fields = sd.fieldNames;
    const rowFieldNames = o.rows;
    const colFieldNames = o.columns ?? [];
    const dataFields = o.data;
    const style = o.style ?? "PivotStyleLight16";
    const location = o.location ?? "A3";
    const name = o.name ?? "PivotTable1";

    // Compute field indices
    const rowFieldIndices = rowFieldNames.map((n) => fields.indexOf(n));
    const colFieldIndices = colFieldNames.map((n) => fields.indexOf(n));
    const dataFieldIndices = dataFields.map((df) => fields.indexOf(df.field));
    const pageFieldNames = o.pages ?? [];
    const pageFieldIndices = pageFieldNames.map((n) => fields.indexOf(n));

    // Build pivotFields
    const pivotFieldsXml = this.buildPivotFields(
      rowFieldIndices,
      colFieldIndices,
      dataFieldIndices,
      pageFieldIndices,
    );

    // Build pageFields (if any)
    const pageFieldsXml = this.buildPageFields(pageFieldIndices);

    // Build rowFields + rowItems
    const rowFieldsXml = this.buildRowFields(rowFieldIndices);
    const rowItemsXml = this.buildRowItems(rowFieldIndices);

    // Build colFields + colItems
    const colFieldsXml = this.buildColFields(colFieldIndices);
    const colItemsXml = this.buildColItems(colFieldIndices, dataFields);

    // Build dataFields
    const dataFieldsXml = this.buildDataFields(dataFields, dataFieldIndices);

    // Compute location ref
    const locationRef = this.computeLocationRef(
      location,
      rowFieldIndices,
      colFieldIndices,
      dataFields,
    );

    const p: string[] = [];
    const defAttrs: string[] = [
      `name="${escapeXml(name)}"`,
      `cacheId="${this.cacheId}"`,
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
    // Optional pivot table definition attributes
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

    // pivotFields
    p.push(pivotFieldsXml);

    // pivotHierarchies (after pivotFields per XSD sequence)
    if (o.pivotHierarchies && o.pivotHierarchies.length > 0) {
      p.push(this.buildPivotHierarchies(o.pivotHierarchies));
    }

    // pageFields (before rowFields per XSD sequence)
    if (pageFieldIndices.length > 0) {
      p.push(pageFieldsXml);
    }

    // rowFields
    p.push(rowFieldsXml);

    // rowItems
    p.push(rowItemsXml);

    // colFields (only when column fields exist)
    if (colFieldIndices.length > 0) {
      p.push(colFieldsXml);
    }

    // colItems (always present)
    p.push(colItemsXml);

    // dataFields
    if (dataFields.length > 0) {
      p.push(dataFieldsXml);
    }

    // style
    p.push(
      `<pivotTableStyleInfo name="${escapeXml(style)}" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>`,
    );

    // filters (after pivotTableStyleInfo per XSD sequence)
    if (o.filters && o.filters.length > 0) {
      p.push(this.buildFilters(o.filters));
    }

    // calculatedItems (after filters per XSD sequence)
    if (o.calculatedItems && o.calculatedItems.length > 0) {
      p.push(this.buildCalculatedItems(o.calculatedItems));
    }

    // calculatedMembers
    if (o.calculatedMembers && o.calculatedMembers.length > 0) {
      p.push(this.buildCalculatedMembers(o.calculatedMembers));
    }

    // formats (CT_Formats, after dataFields per XSD sequence)
    if (o.formats && o.formats.length > 0) {
      const fmtParts: string[] = [`<formats count="${o.formats.length}">`];
      for (const fmt of o.formats) {
        const fmtAttrs: string[] = [];
        if (fmt.action && fmt.action !== "formatting") fmtAttrs.push(`action="${fmt.action}"`);
        if (fmt.dxfId !== undefined) fmtAttrs.push(`dxfId="${fmt.dxfId}"`);
        fmtParts.push(
          `<format${fmtAttrs.length ? " " + fmtAttrs.join(" ") : ""}>${this.buildPivotAreaXml(fmt.pivotArea)}</format>`,
        );
      }
      fmtParts.push("</formats>");
      p.push(fmtParts.join(""));
    }

    // conditionalFormats (pivot-specific)
    if (o.pivotConditionalFormats && o.pivotConditionalFormats.length > 0) {
      p.push(this.buildPivotConditionalFormats(o.pivotConditionalFormats));
    }

    // chartFormats
    if (o.chartFormats && o.chartFormats.length > 0) {
      p.push(this.buildChartFormats(o.chartFormats));
    }

    // rowHierarchiesUsage (CT_RowHierarchiesUsage)
    if (o.rowHierarchiesUsage && o.rowHierarchiesUsage.length > 0) {
      const rhu = o.rowHierarchiesUsage;
      p.push(
        `<rowHierarchiesUsage count="${rhu.length}">${rhu.map((h) => `<rowHierarchyUsage hierarchyUsage="${h.hierarchyUsage}"/>`).join("")}</rowHierarchiesUsage>`,
      );
    }

    // colHierarchiesUsage (CT_ColHierarchiesUsage)
    if (o.colHierarchiesUsage && o.colHierarchiesUsage.length > 0) {
      const chu = o.colHierarchiesUsage;
      p.push(
        `<colHierarchiesUsage count="${chu.length}">${chu.map((h) => `<colHierarchyUsage hierarchyUsage="${h.hierarchyUsage}"/>`).join("")}</colHierarchiesUsage>`,
      );
    }

    p.push("</pivotTableDefinition>");
    return p.join("");
  }

  private buildFieldOverrideAttrs(fo: PivotFieldOverrideOptions): string {
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

  private buildPivotFields(
    rowIndices: readonly number[],
    colIndices: readonly number[],
    dataIndices: readonly number[],
    pageIndices: readonly number[],
  ): string {
    const fields = this.sourceData.fieldNames;
    const o = this.options;
    const parts: string[] = [`<pivotFields count="${fields.length}">`];

    for (let i = 0; i < fields.length; i++) {
      const isRow = rowIndices.includes(i);
      const isCol = colIndices.includes(i);
      const isData = dataIndices.includes(i);
      const isPage = pageIndices.includes(i);
      const override = o.fieldOverrides?.find((fo) => fo.field === fields[i]);
      const extraAttrs = override ? this.buildFieldOverrideAttrs(override) : "";

      if (isData) {
        const dataFieldIdx = dataIndices.indexOf(i);
        const df = o.data[dataFieldIdx];
        const dfAttrs: string[] = ['dataField="1"', 'showAll="0"'];
        if (extraAttrs) dfAttrs.push(extraAttrs);
        if (df?.showDataAs) dfAttrs.push(`showDataAs="${df.showDataAs}"`);
        if (df?.baseField !== undefined) dfAttrs.push(`baseField="${df.baseField}"`);
        if (df?.baseItem !== undefined) dfAttrs.push(`baseItem="${df.baseItem}"`);
        // autoSortScope for data fields
        if (o.autoSortScope || (df?.sortByTupleItems && df.sortByTupleItems.length > 0)) {
          const scopeChildren: string[] = [];
          if (o.autoSortScope) {
            scopeChildren.push(this.buildPivotAreaXml(o.autoSortScope));
          }
          if (df?.sortByTupleItems && df.sortByTupleItems.length > 0) {
            const tplXml = df.sortByTupleItems.map((v) => `<tpl><x v="${v}"/></tpl>`).join("");
            scopeChildren.push(`<sortByTuple>${tplXml}</sortByTuple>`);
          }
          parts.push(
            `<pivotField ${dfAttrs.join(" ")}><autoSortScope>${scopeChildren.join("")}</autoSortScope></pivotField>`,
          );
        } else {
          parts.push(`<pivotField ${dfAttrs.join(" ")}/>`);
        }
      } else if (isRow) {
        const uniqueVals = collectUniqueValues(this.sourceData.records, i);
        const rAttrs = extraAttrs
          ? ` axis="axisRow" showAll="0" ${extraAttrs}`
          : ' axis="axisRow" showAll="0"';
        parts.push(`<pivotField${rAttrs}>`);
        parts.push(`<items count="${uniqueVals.length + 1}">`);
        for (let j = 0; j < uniqueVals.length; j++) {
          parts.push(`<item x="${j}"/>`);
        }
        parts.push(`<item t="default"${override?.defaultItemSd === false ? ' sd="0"' : ""}/>`);
        parts.push("</items></pivotField>");
      } else if (isCol) {
        const uniqueVals = collectUniqueValues(this.sourceData.records, i);
        const cAttrs = extraAttrs
          ? ` axis="axisCol" showAll="0" ${extraAttrs}`
          : ' axis="axisCol" showAll="0"';
        parts.push(`<pivotField${cAttrs}>`);
        parts.push(`<items count="${uniqueVals.length + 1}">`);
        for (let j = 0; j < uniqueVals.length; j++) {
          parts.push(`<item x="${j}"/>`);
        }
        parts.push(`<item t="default"${override?.defaultItemSd === false ? ' sd="0"' : ""}/>`);
        parts.push("</items></pivotField>");
      } else if (isPage) {
        const uniqueVals = collectUniqueValues(this.sourceData.records, i);
        const pAttrs = extraAttrs
          ? ` axis="axisPage" showAll="0" ${extraAttrs}`
          : ' axis="axisPage" showAll="0"';
        parts.push(`<pivotField${pAttrs}>`);
        parts.push(`<items count="${uniqueVals.length + 1}">`);
        for (let j = 0; j < uniqueVals.length; j++) {
          parts.push(`<item x="${j}"/>`);
        }
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

  private buildPageFields(pageIndices: readonly number[]): string {
    if (pageIndices.length === 0) return "";
    const o = this.options;
    const parts: string[] = [`<pageFields count="${pageIndices.length}">`];
    for (let i = 0; i < pageIndices.length; i++) {
      const cap = o.pageCaptions?.[i];
      const capAttr = cap ? ` cap="${escapeXml(cap)}"` : "";
      parts.push(`<pageField fld="${pageIndices[i]}" hier="${i}"${capAttr}/>`);
    }
    parts.push("</pageFields>");
    return parts.join("");
  }

  private buildRowFields(rowIndices: readonly number[]): string {
    if (rowIndices.length === 0) return '<rowFields count="0"/>';
    const parts: string[] = [`<rowFields count="${rowIndices.length}">`];
    for (const idx of rowIndices) {
      parts.push(`<field x="${idx}"/>`);
    }
    parts.push("</rowFields>");
    return parts.join("");
  }

  private buildRowItems(rowIndices: readonly number[]): string {
    if (rowIndices.length === 0) return '<rowItems count="1"><i/></rowItems>';

    const allUniqueCounts: number[] = [];
    for (const idx of rowIndices) {
      const vals = collectUniqueValues(this.sourceData.records, idx);
      allUniqueCounts.push(vals.length);
    }

    // Simple case: single row field
    if (rowIndices.length === 1) {
      const count = allUniqueCounts[0];
      const parts: string[] = [`<rowItems count="${count + 1}">`];
      for (let i = 0; i < count; i++) {
        parts.push(`<i><x v="${i}"/></i>`);
      }
      parts.push(`<i t="grand"><x/></i>`);
      parts.push("</rowItems>");
      return parts.join("");
    }

    // Multi-field: build cartesian product indices
    const combos = cartesianOfCounts(allUniqueCounts);
    const rowItems: string[] = [];
    for (const combo of combos) {
      const xParts = combo.map((v) => `<x v="${v}"/>`).join("");
      rowItems.push(`<i>${xParts}</i>`);
    }
    // Grand total
    rowItems.push(`<i t="grand">${rowIndices.map(() => "<x/>").join("")}</i>`);

    return `<rowItems count="${rowItems.length}">${rowItems.join("")}</rowItems>`;
  }

  private buildColFields(colIndices: readonly number[]): string {
    if (colIndices.length === 0) return '<colFields count="0"/>';
    const parts: string[] = [`<colFields count="${colIndices.length}">`];
    for (const idx of colIndices) {
      parts.push(`<field x="${idx}"/>`);
    }
    parts.push("</colFields>");
    return parts.join("");
  }

  private buildColItems(
    colIndices: readonly number[],
    dataFields: readonly PivotDataField[],
  ): string {
    if (colIndices.length > 0) {
      const allUniqueCounts: number[] = [];
      for (const idx of colIndices) {
        const vals = collectUniqueValues(this.sourceData.records, idx);
        allUniqueCounts.push(vals.length);
      }

      const combos = cartesianOfCounts(allUniqueCounts);
      const items: string[] = [];
      for (const combo of combos) {
        const xParts = combo.map((v) => `<x v="${v}"/>`).join("");
        items.push(`<i>${xParts}</i>`);
      }
      items.push(`<i t="grand">${colIndices.map(() => "<x/>").join("")}</i>`);
      return `<colItems count="${items.length}">${items.join("")}</colItems>`;
    }

    // No column fields: if multiple data fields, each becomes a column
    if (dataFields.length > 1) {
      const items = dataFields.map((_, i) => `<i><x v="${i}"/></i>`);
      return `<colItems count="${items.length}">${items.join("")}</colItems>`;
    }

    // Single or no data fields, no column fields
    return '<colItems count="1"><i/></colItems>';
  }

  private buildDataFields(
    dataFields: readonly PivotDataField[],
    dataFieldIndices: readonly number[],
  ): string {
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

  private computeLocationRef(
    location: string,
    rowFieldIndices: readonly number[],
    colFieldIndices: readonly number[],
    dataFields: readonly PivotDataField[],
  ): string {
    const startCell = location.split(":")[0];
    const match = startCell.match(/^([A-Z]+)(\d+)$/);
    if (!match) return location;

    const startCol = match[1];
    const startRow = parseInt(match[2], 10);

    // Count rows: unique row values + header + grand total
    let rowCount = 1; // header
    if (rowFieldIndices.length > 0) {
      rowCount += collectUniqueValues(this.sourceData.records, rowFieldIndices[0]).length;
    }
    rowCount += 1; // grand total

    // Count columns: row fields + column values or data fields
    let colCount = Math.max(rowFieldIndices.length, 1);
    if (colFieldIndices.length > 0) {
      colCount += collectUniqueValues(this.sourceData.records, colFieldIndices[0]).length;
    } else if (dataFields.length > 1) {
      colCount += dataFields.length - 1;
    }
    colCount += 1; // grand total column

    const endCol = colIndexToLetter(letterToColIndex(startCol) + colCount - 1);
    const endRow = startRow + rowCount - 1;

    return `${startCol}${startRow}:${endCol}${endRow}`;
  }

  private buildFilters(filters: readonly PivotFilterOptions[]): string {
    const parts: string[] = [`<filters count="${filters.length}">`];
    for (const f of filters) {
      const fAttrs: Record<string, string | number | boolean | undefined> = {
        fld: f.fld,
        type: f.type,
        id: f.id,
      };
      if (f.mpFld !== undefined) fAttrs.mpFld = f.mpFld;
      if (f.evalOrder !== undefined) fAttrs.evalOrder = f.evalOrder;
      if (f.iMeasureHier !== undefined) fAttrs.iMeasureHier = f.iMeasureHier;
      if (f.iMeasureFld !== undefined) fAttrs.iMeasureFld = f.iMeasureFld;
      if (f.name !== undefined) fAttrs.name = f.name;
      if (f.description !== undefined) fAttrs.description = f.description;
      if (f.stringValue1 !== undefined) fAttrs.stringValue1 = f.stringValue1;
      if (f.stringValue2 !== undefined) fAttrs.stringValue2 = f.stringValue2;
      parts.push(`<filter${attrs(fAttrs)}><autoFilter/></filter>`);
    }
    parts.push("</filters>");
    return parts.join("");
  }

  private buildPivotHierarchies(hierarchies: readonly PivotHierarchyOptions[]): string {
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
      const mpsXml = h.memberProperties
        ? `<mps count="${h.memberProperties.length}">${h.memberProperties
            .map((mp) => {
              const mpAttrs: string[] = [`field="${mp.field}"`];
              if (mp.name !== undefined) mpAttrs.push(`name="${escapeXml(mp.name)}"`);
              if (mp.showCell) mpAttrs.push('showCell="1"');
              if (mp.showTip) mpAttrs.push('showTip="1"');
              if (mp.showAsCaption) mpAttrs.push('showAsCaption="1"');
              if (mp.nameLen !== undefined) mpAttrs.push(`nameLen="${mp.nameLen}"`);
              if (mp.pPos !== undefined) mpAttrs.push(`pPos="${mp.pPos}"`);
              if (mp.pLen !== undefined) mpAttrs.push(`pLen="${mp.pLen}"`);
              return `<mp ${mpAttrs.join(" ")}/>`;
            })
            .join("")}</mps>`
        : "";
      const membersXml = h.members
        ? `<members count="${h.members.length}">${h.members
            .map((m) => {
              const levelAttr = m.level !== undefined ? ` level="${m.level}"` : "";
              return `<member name="${escapeXml(m.name)}"${levelAttr}/>`;
            })
            .join("")}</members>`
        : "";
      const inner = mpsXml + membersXml;
      if (inner) {
        parts.push(`<pivotHierarchy ${hAttrs.join(" ")}>${inner}</pivotHierarchy>`);
      } else {
        parts.push(`<pivotHierarchy ${hAttrs.join(" ")}/>`);
      }
    }
    parts.push("</pivotHierarchies>");
    return parts.join("");
  }

  private buildCalculatedItems(items: readonly CalculatedItemOptions[]): string {
    const parts: string[] = [`<calculatedItems count="${items.length}">`];
    for (const item of items) {
      const ciAttrs: string[] = [];
      if (item.field !== undefined) ciAttrs.push(`field="${item.field}"`);
      if (item.formula) ciAttrs.push(`formula="${escapeXml(item.formula)}"`);
      const inner = item.pivotArea ? this.buildPivotAreaXml(item.pivotArea) : "";
      if (inner) {
        parts.push(`<calculatedItem ${ciAttrs.join(" ")}>${inner}</calculatedItem>`);
      } else {
        parts.push(`<calculatedItem ${ciAttrs.join(" ")}/>`);
      }
    }
    parts.push("</calculatedItems>");
    return parts.join("");
  }

  private buildCalculatedMembers(members: readonly CalculatedMemberOptions[]): string {
    const parts: string[] = [`<calculatedMembers count="${members.length}">`];
    for (const m of members) {
      const mAttrs: string[] = [`name="${escapeXml(m.name)}"`, `mdx="${escapeXml(m.mdx)}"`];
      if (m.memberName) mAttrs.push(`memberName="${escapeXml(m.memberName)}"`);
      if (m.hierarchy) mAttrs.push(`hierarchy="${escapeXml(m.hierarchy)}"`);
      if (m.parent) mAttrs.push(`parent="${escapeXml(m.parent)}"`);
      if (m.solveOrder !== undefined) mAttrs.push(`solveOrder="${m.solveOrder}"`);
      if (m.set) mAttrs.push('set="1"');
      parts.push(`<calculatedMember ${mAttrs.join(" ")}/>`);
    }
    parts.push("</calculatedMembers>");
    return parts.join("");
  }

  private buildPivotConditionalFormats(formats: readonly PivotConditionalFormatOptions[]): string {
    const parts: string[] = [`<conditionalFormats count="${formats.length}">`];
    for (const cf of formats) {
      const cfAttrs: string[] = [`priority="${cf.priority}"`];
      if (cf.scope && cf.scope !== "selection") cfAttrs.push(`scope="${cf.scope}"`);
      if (cf.type && cf.type !== "none") cfAttrs.push(`type="${cf.type}"`);
      const areasXml = cf.pivotAreas?.map((a) => this.buildPivotAreaXml(a)).join("") ?? "";
      const pivotAreasXml = areasXml ? `<pivotAreas>${areasXml}</pivotAreas>` : "";
      parts.push(`<conditionalFormat ${cfAttrs.join(" ")}>${pivotAreasXml}</conditionalFormat>`);
    }
    parts.push("</conditionalFormats>");
    return parts.join("");
  }

  private buildChartFormats(formats: readonly ChartFormatOptions[]): string {
    const parts: string[] = [`<chartFormats count="${formats.length}">`];
    for (const cf of formats) {
      const cfAttrs: string[] = [`chart="${cf.chart}"`, `format="${cf.format}"`];
      if (cf.series) cfAttrs.push('series="1"');
      const areaXml = cf.pivotArea ? this.buildPivotAreaXml(cf.pivotArea) : "";
      if (areaXml) {
        parts.push(`<chartFormat ${cfAttrs.join(" ")}>${areaXml}</chartFormat>`);
      } else {
        parts.push(`<chartFormat ${cfAttrs.join(" ")}/>`);
      }
    }
    parts.push("</chartFormats>");
    return parts.join("");
  }

  private buildPivotAreaXml(area: PivotAreaOptions): string {
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
    const refsXml = area.references ? this.buildPivotAreaReferences(area.references) : "";
    if (refsXml) {
      return `<pivotArea ${aAttrs.join(" ")}>${refsXml}</pivotArea>`;
    }
    return `<pivotArea ${aAttrs.join(" ")}/>`;
  }

  private buildPivotAreaReferences(
    refs: readonly import("./pivot-utils").PivotAreaReferenceOptions[],
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
      if (ref.sumSubtotal) rAttrs.push('sumSubtotal="1"');
      if (ref.countASubtotal) rAttrs.push('countASubtotal="1"');
      if (ref.avgSubtotal) rAttrs.push('avgSubtotal="1"');
      if (ref.maxSubtotal) rAttrs.push('maxSubtotal="1"');
      if (ref.minSubtotal) rAttrs.push('minSubtotal="1"');
      if (ref.countSubtotal) rAttrs.push('countSubtotal="1"');
      if (ref.productSubtotal) rAttrs.push('productSubtotal="1"');
      if (ref.stdDevPSubtotal) rAttrs.push('stdDevPSubtotal="1"');
      if (ref.stdDevSubtotal) rAttrs.push('stdDevSubtotal="1"');
      if (ref.varPSubtotal) rAttrs.push('varPSubtotal="1"');
      if (ref.varSubtotal) rAttrs.push('varSubtotal="1"');
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
}

function letterToColIndex(letters: string): number {
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
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

/** Cartesian product from counts: e.g. [2, 3] → [[0,0],[0,1],[0,2],[1,0],[1,1],[1,2]] */
function cartesianOfCounts(counts: readonly number[]): number[][] {
  if (counts.length === 0) return [[]];
  let result: number[][] = [[]];
  for (const count of counts) {
    const next: number[][] = [];
    for (const prefix of result) {
      for (let i = 0; i < count; i++) {
        next.push([...prefix, i]);
      }
    }
    result = next;
  }
  return result;
}
