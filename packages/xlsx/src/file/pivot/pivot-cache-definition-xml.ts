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

import type {
  PivotSourceData,
  OlapPrOptions,
  CacheHierarchyOptions,
  KpiOptions,
  MeasureGroupOptions,
  SetOptions,
  ServerFormatOptions,
  PivotDimensionOptions,
  FieldGroupOptions,
} from "./pivot-utils";
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
  /** Cache hierarchies (CT_CacheHierarchies) */
  readonly cacheHierarchies?: readonly CacheHierarchyOptions[];
  /** KPIs (CT_PCDKPIs) */
  readonly kpis?: readonly KpiOptions[];
  /** Measure groups (CT_MeasureGroups) */
  readonly measureGroups?: readonly MeasureGroupOptions[];
  /** Dimensions (CT_Dimensions) */
  readonly dimensions?: readonly PivotDimensionOptions[];
  /** Sets (CT_Sets in tupleCache) */
  readonly sets?: readonly SetOptions[];
  /** Server formats (CT_ServerFormats) */
  readonly serverFormats?: readonly ServerFormatOptions[];
  /** Field groups per field index (CT_FieldGroup inside cacheField) */
  readonly fieldGroups?: ReadonlyMap<number, FieldGroupOptions>;
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
      if (this.olapPr.localRefresh) olAttrs.push(' localRefresh="1"');
      if (this.olapPr.serverFill === false) olAttrs.push(' serverFill="0"');
      if (this.olapPr.serverNumberFormat === false) olAttrs.push(' serverNumberFormat="0"');
      if (this.olapPr.serverFont === false) olAttrs.push(' serverFont="0"');
      if (this.olapPr.serverFontColor === false) olAttrs.push(' serverFontColor="0"');
      if (olAttrs.length > 0) {
        p.push(`<olapPr${olAttrs.join("")}/>`);
      }
    }

    // cacheHierarchies (optional)
    if (this.cacheDefOpts?.cacheHierarchies && this.cacheDefOpts.cacheHierarchies.length > 0) {
      const chs = this.cacheDefOpts.cacheHierarchies;
      p.push(`<cacheHierarchies count="${chs.length}">`);
      for (const ch of chs) {
        const chAttrs: string[] = [
          `uniqueName="${escapeXml(ch.uniqueName)}"`,
          `count="${ch.count}"`,
        ];
        if (ch.caption) chAttrs.push(`caption="${escapeXml(ch.caption)}"`);
        if (ch.measure) chAttrs.push('measure="1"');
        if (ch.set) chAttrs.push('set="1"');
        if (ch.parentSet !== undefined) chAttrs.push(`parentSet="${ch.parentSet}"`);
        if (ch.iconSet !== undefined && ch.iconSet !== 0) chAttrs.push(`iconSet="${ch.iconSet}"`);
        if (ch.attribute) chAttrs.push('attribute="1"');
        if (ch.time) chAttrs.push('time="1"');
        if (ch.keyAttribute) chAttrs.push('keyAttribute="1"');
        if (ch.defaultMemberUniqueName)
          chAttrs.push(`defaultMemberUniqueName="${escapeXml(ch.defaultMemberUniqueName)}"`);
        if (ch.allUniqueName) chAttrs.push(`allUniqueName="${escapeXml(ch.allUniqueName)}"`);
        if (ch.allCaption) chAttrs.push(`allCaption="${escapeXml(ch.allCaption)}"`);
        if (ch.dimensionUniqueName)
          chAttrs.push(`dimensionUniqueName="${escapeXml(ch.dimensionUniqueName)}"`);
        if (ch.displayFolder) chAttrs.push(`displayFolder="${escapeXml(ch.displayFolder)}"`);
        if (ch.measureGroup) chAttrs.push(`measureGroup="${escapeXml(ch.measureGroup)}"`);
        if (ch.measures) chAttrs.push('measures="1"');
        if (ch.oneField) chAttrs.push('oneField="1"');
        if (ch.hidden) chAttrs.push('hidden="1"');
        p.push(`<cacheHierarchy ${chAttrs.join(" ")}/>`);
      }
      p.push("</cacheHierarchies>");
    }

    // kpis (optional)
    if (this.cacheDefOpts?.kpis && this.cacheDefOpts.kpis.length > 0) {
      const kpis = this.cacheDefOpts.kpis;
      p.push(`<kpis count="${kpis.length}">`);
      for (const k of kpis) {
        const kAttrs: string[] = [
          `uniqueName="${escapeXml(k.uniqueName)}"`,
          `value="${escapeXml(k.value)}"`,
        ];
        if (k.caption) kAttrs.push(`caption="${escapeXml(k.caption)}"`);
        if (k.displayFolder) kAttrs.push(`displayFolder="${escapeXml(k.displayFolder)}"`);
        if (k.measureGroup) kAttrs.push(`measureGroup="${escapeXml(k.measureGroup)}"`);
        if (k.parent) kAttrs.push(`parent="${escapeXml(k.parent)}"`);
        if (k.goal) kAttrs.push(`goal="${escapeXml(k.goal)}"`);
        if (k.status) kAttrs.push(`status="${escapeXml(k.status)}"`);
        if (k.trend) kAttrs.push(`trend="${escapeXml(k.trend)}"`);
        if (k.weight) kAttrs.push(`weight="${escapeXml(k.weight)}"`);
        if (k.time) kAttrs.push(`time="${escapeXml(k.time)}"`);
        p.push(`<kpi ${kAttrs.join(" ")}/>`);
      }
      p.push("</kpis>");
    }

    // measureGroups (optional)
    if (this.cacheDefOpts?.measureGroups && this.cacheDefOpts.measureGroups.length > 0) {
      const mgs = this.cacheDefOpts.measureGroups;
      p.push(`<measureGroups count="${mgs.length}">`);
      for (const mg of mgs) {
        p.push(`<measureGroup name="${escapeXml(mg.name)}" caption="${escapeXml(mg.caption)}"/>`);
      }
      p.push("</measureGroups>");
    }

    // dimensions (optional)
    if (this.cacheDefOpts?.dimensions && this.cacheDefOpts.dimensions.length > 0) {
      const dims = this.cacheDefOpts.dimensions;
      p.push(`<dimensions count="${dims.length}">`);
      for (const d of dims) {
        const dAttrs: string[] = [
          `name="${escapeXml(d.name)}"`,
          `uniqueName="${escapeXml(d.uniqueName)}"`,
          `caption="${escapeXml(d.caption)}"`,
        ];
        if (d.measure) dAttrs.push('measure="1"');
        p.push(`<dimension ${dAttrs.join(" ")}/>`); // XSD uses "dimension" not "dimensions" for individual items
      }
      p.push("</dimensions>");
    }

    // tupleCache with sets and serverFormats (optional)
    const cd = this.cacheDefOpts;
    const hasSets = cd?.sets && cd.sets.length > 0;
    const hasServerFormats = cd?.serverFormats && cd.serverFormats.length > 0;
    if (hasSets || hasServerFormats) {
      p.push("<tupleCache>");
      if (hasSets) {
        p.push(`<sets count="${cd!.sets!.length}">`);
        for (const s of cd!.sets!) {
          const sAttrs: string[] = [
            `maxRank="${s.maxRank}"`,
            `setDefinition="${escapeXml(s.setDefinition)}"`,
          ];
          if (s.count !== undefined) sAttrs.push(`count="${s.count}"`);
          if (s.sortType && s.sortType !== "none") sAttrs.push(`sortType="${s.sortType}"`);
          if (s.queryFailed) sAttrs.push('queryFailed="1"');
          p.push(`<set ${sAttrs.join(" ")}/>`);
        }
        p.push("</sets>");
      }
      if (hasServerFormats) {
        p.push(`<serverFormats count="${cd!.serverFormats!.length}">`);
        for (const sf of cd!.serverFormats!) {
          const sfAttrs: string[] = [];
          if (sf.culture) sfAttrs.push(`culture="${escapeXml(sf.culture)}"`);
          if (sf.format) sfAttrs.push(`format="${escapeXml(sf.format)}"`);
          p.push(`<serverFormat ${sfAttrs.join(" ")}/>`);
        }
        p.push("</serverFormats>");
      }
      p.push("</tupleCache>");
    }

    p.push("</pivotCacheDefinition>");

    return p.join("");
  }
}
