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
import type { PivotTableOptions, PivotDataField, PivotFilterOptions } from "./pivot-utils";
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

    // Build pivotFields
    const pivotFieldsXml = this.buildPivotFields(
      rowFieldIndices,
      colFieldIndices,
      dataFieldIndices,
    );

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
    p.push(
      `<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` +
        ` name="${escapeXml(name)}" cacheId="${this.cacheId}"` +
        ` dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" createdVersion="6"` +
        ` applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0"` +
        ` applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1"` +
        ` autoFormatId="0" useAutoFormatting="1" itemPrintTitles="1" indent="0"` +
        ` outline="1" outlineData="1" compact="1" compactData="1" rowGrandTotals="1" colGrandTotals="1">`,
    );

    // location
    p.push(
      `<location ref="${escapeXml(locationRef)}" firstHeaderRow="1" firstDataRow="${colFieldIndices.length + 1}" firstDataCol="${rowFieldIndices.length}"/>`,
    );

    // pivotFields
    p.push(pivotFieldsXml);

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

    p.push("</pivotTableDefinition>");
    return p.join("");
  }

  private buildPivotFields(
    rowIndices: readonly number[],
    colIndices: readonly number[],
    dataIndices: readonly number[],
  ): string {
    const fields = this.sourceData.fieldNames;
    const parts: string[] = [`<pivotFields count="${fields.length}">`];

    for (let i = 0; i < fields.length; i++) {
      const isRow = rowIndices.includes(i);
      const isCol = colIndices.includes(i);
      const isData = dataIndices.includes(i);

      if (isData) {
        parts.push(`<pivotField dataField="1" showAll="0"/>`);
      } else if (isRow) {
        const uniqueVals = collectUniqueValues(this.sourceData.records, i);
        parts.push(`<pivotField axis="axisRow" showAll="0">`);
        parts.push(`<items count="${uniqueVals.length + 1}">`);
        for (let j = 0; j < uniqueVals.length; j++) {
          parts.push(`<item x="${j}"/>`);
        }
        parts.push(`<item t="default"/>`);
        parts.push("</items></pivotField>");
      } else if (isCol) {
        const uniqueVals = collectUniqueValues(this.sourceData.records, i);
        parts.push(`<pivotField axis="axisCol" showAll="0">`);
        parts.push(`<items count="${uniqueVals.length + 1}">`);
        for (let j = 0; j < uniqueVals.length; j++) {
          parts.push(`<item x="${j}"/>`);
        }
        parts.push(`<item t="default"/>`);
        parts.push("</items></pivotField>");
      } else {
        parts.push(`<pivotField showAll="0"/>`);
      }
    }

    parts.push("</pivotFields>");
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
      parts.push(
        `<dataField name="${escapeXml(name)}" fld="${dataFieldIndices[i]}" subtotal="${subtotal}"/>`,
      );
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
