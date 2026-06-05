/**
 * PivotCacheDefinition XML generator.
 *
 * Generates xl/pivotCache/pivotCacheDefinition{N}.xml.
 * Follows CT_PivotCacheDefinition from sml.xsd.
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { escapeXml } from "@office-open/xml";

import type { PivotSourceData, OlapPrOptions } from "./pivot-utils";
import { collectUniqueValues, isNumericField } from "./pivot-utils";

export interface PivotCacheDefinitionOptions {
  /** Cache is invalid (CT_PivotCacheDefinition @invalid) */
  readonly invalid?: boolean;
  /** Save data with cache (CT_PivotCacheDefinition @saveData) */
  readonly saveData?: boolean;
  /** Optimize memory usage (CT_PivotCacheDefinition @optimizeMemory) */
  readonly optimizeMemory?: boolean;
  /** Enable refresh (CT_PivotCacheDefinition @enableRefresh) */
  readonly enableRefresh?: boolean;
  /** User who last refreshed */
  readonly refreshedBy?: string;
  /** Date of last refresh (decimal) */
  readonly refreshedDate?: number;
  /** Date of last refresh (ISO 8601) */
  readonly refreshedDateIso?: string;
  /** Background query (CT_PivotCacheDefinition @backgroundQuery) */
  readonly backgroundQuery?: boolean;
  /** Missing items limit */
  readonly missingItemsLimit?: number;
  /** Upgrade on refresh */
  readonly upgradeOnRefresh?: boolean;
  /** Support subquery */
  readonly supportSubquery?: boolean;
  /** Support advanced drill */
  readonly supportAdvancedDrill?: boolean;
}

export class PivotCacheDefinitionXml extends BaseXmlComponent {
  private readonly sourceRef: string;
  private readonly sourceSheet: string;
  private readonly sourceData: PivotSourceData;
  private readonly recordsRid: string;
  private readonly olapPr?: OlapPrOptions;
  private readonly cacheDefOpts?: PivotCacheDefinitionOptions;

  public constructor(
    _cacheIdx: number,
    sourceRef: string,
    sourceSheet: string,
    sourceData: PivotSourceData,
    recordsRid: string,
    olapPr?: OlapPrOptions,
    cacheDefOpts?: PivotCacheDefinitionOptions,
  ) {
    super("pivotCacheDefinition");
    this.sourceRef = sourceRef;
    this.sourceSheet = sourceSheet;
    this.sourceData = sourceData;
    this.recordsRid = recordsRid;
    this.olapPr = olapPr;
    this.cacheDefOpts = cacheDefOpts;
  }

  public override toXml(_context: Context): string {
    const p: string[] = [];

    // Root element attributes
    const rootAttrs: string[] = [
      'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
      `r:id="${escapeXml(this.recordsRid)}"`,
      `recordCount="${this.sourceData.records.length}"`,
      'createdVersion="6"',
      'refreshedVersion="6"',
      'minRefreshableVersion="3"',
    ];

    // Optional cache definition attributes on root element
    if (this.cacheDefOpts) {
      const cd = this.cacheDefOpts;
      if (cd.invalid) rootAttrs.push('invalid="1"');
      if (cd.saveData === false) rootAttrs.push('saveData="0"');
      if (cd.optimizeMemory) rootAttrs.push('optimizeMemory="1"');
      if (cd.enableRefresh === false) rootAttrs.push('enableRefresh="0"');
      if (cd.refreshedBy) rootAttrs.push(`refreshedBy="${escapeXml(cd.refreshedBy)}"`);
      if (cd.refreshedDate !== undefined) rootAttrs.push(`refreshedDate="${cd.refreshedDate}"`);
      if (cd.refreshedDateIso)
        rootAttrs.push(`refreshedDateIso="${escapeXml(cd.refreshedDateIso)}"`);
      if (cd.backgroundQuery) rootAttrs.push('backgroundQuery="1"');
      if (cd.missingItemsLimit !== undefined)
        rootAttrs.push(`missingItemsLimit="${cd.missingItemsLimit}"`);
      if (cd.upgradeOnRefresh) rootAttrs.push('upgradeOnRefresh="1"');
      if (cd.supportSubquery) rootAttrs.push('supportSubquery="1"');
      if (cd.supportAdvancedDrill) rootAttrs.push('supportAdvancedDrill="1"');
    }

    p.push(`<pivotCacheDefinition ${rootAttrs.join(" ")}>`);

    // cacheSource
    p.push(
      `<cacheSource type="worksheet">` +
        `<worksheetSource ref="${escapeXml(this.sourceRef)}" sheet="${escapeXml(this.sourceSheet)}"/>` +
        `</cacheSource>`,
    );

    // cacheFields
    const fields = this.sourceData.fieldNames;
    p.push(`<cacheFields count="${fields.length}">`);

    for (let i = 0; i < fields.length; i++) {
      const fieldName = fields[i];
      const numeric = isNumericField(this.sourceData.records, i);
      const uniqueVals = collectUniqueValues(this.sourceData.records, i);

      if (numeric) {
        // Numeric field: sharedItems with type metadata only, no inline values
        let min = Infinity;
        let max = -Infinity;
        for (const row of this.sourceData.records) {
          const v = row[i];
          if (typeof v === "number") {
            if (v < min) min = v;
            if (v > max) max = v;
          }
        }
        if (!isFinite(min)) {
          min = 0;
          max = 0;
        }
        const allInteger = this.sourceData.records.every(
          (row) => typeof row[i] === "number" && Number.isInteger(row[i]),
        );
        p.push(
          `<cacheField name="${escapeXml(fieldName)}" numFmtId="0">` +
            `<sharedItems containsSemiMixedTypes="0" containsString="0"` +
            ` containsNumber="1" containsInteger="${allInteger ? "1" : "0"}"` +
            ` minValue="${min}" maxValue="${max}" count="${uniqueVals.length}"/>` +
            `</cacheField>`,
        );
      } else {
        // String/categorical field: list unique values in sharedItems
        p.push(
          `<cacheField name="${escapeXml(fieldName)}" numFmtId="0"><sharedItems count="${uniqueVals.length}">`,
        );
        for (const v of uniqueVals) {
          p.push(`<s v="${escapeXml(String(v))}"/>`);
        }
        p.push("</sharedItems></cacheField>");
      }
    }

    p.push("</cacheFields>");

    // olapPr (optional)
    if (this.olapPr) {
      const olAttrs: string[] = [];
      if (this.olapPr.local) olAttrs.push(` local="${escapeXml(this.olapPr.local)}"`);
      if (this.olapPr.localConnection)
        olAttrs.push(` localConnection="${escapeXml(this.olapPr.localConnection)}"`);
      if (this.olapPr.sendLocale) olAttrs.push(` sendLocale="1"`);
      if (this.olapPr.rowDrillCount !== undefined)
        olAttrs.push(` rowDrillCount="${this.olapPr.rowDrillCount}"`);
      if (this.olapPr.colDrillCount !== undefined)
        olAttrs.push(` colDrillCount="${this.olapPr.colDrillCount}"`);
      if (olAttrs.length > 0) {
        p.push(`<olapPr${olAttrs.join("")}/>`);
      }
    }

    p.push("</pivotCacheDefinition>");

    return p.join("");
  }
}
