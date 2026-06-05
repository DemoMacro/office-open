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
  ConsolidationOptions,
  TupleCacheEntryOptions,
  QueryCacheEntryOptions,
  MpMapOptions,
  MeasureDimensionMapOptions,
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
  /** Consolidation source (alternative to worksheetSource) */
  readonly consolidation?: ConsolidationOptions;
  /** Tuple cache entries (CT_PCDSDTCEntries) */
  readonly entries?: readonly TupleCacheEntryOptions[];
  /** Query cache (CT_QueryCache in tupleCache) */
  readonly queryCache?: readonly QueryCacheEntryOptions[];
  /** Member property map per cache field (mpMap) */
  readonly mpMaps?: readonly MpMapOptions[];
  /** Measure dimension maps (CT_MeasureDimensionMaps) */
  readonly measureDimensionMaps?: readonly MeasureDimensionMapOptions[];
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

    // cacheSource (worksheetSource or consolidation)
    if (this.cacheDefOpts?.consolidation) {
      const con = this.cacheDefOpts.consolidation;
      const conParts: string[] = ['<cacheSource type="consolidation"><consolidation'];
      if (con.autoPage === false) conParts.push(' autoPage="0"');
      conParts.push(">");
      // pages (optional, max 4)
      if (con.pages && con.pages.length > 0) {
        conParts.push(`<pages count="${con.pages.length}">`);
        for (const pg of con.pages) {
          const pgItems = pg.items ?? [];
          conParts.push(`<page${pgItems.length ? ` count="${pgItems.length}"` : ""}>`);
          for (const pi of pgItems) {
            conParts.push(`<pageItem name="${escapeXml(pi.name)}"/>`);
          }
          conParts.push("</page>");
        }
        conParts.push("</pages>");
      }
      // rangeSets (required)
      conParts.push(`<rangeSets count="${con.rangeSets.length}">`);
      for (const rs of con.rangeSets) {
        const rsAttrs: string[] = [];
        if (rs.i1 !== undefined) rsAttrs.push(`i1="${rs.i1}"`);
        if (rs.i2 !== undefined) rsAttrs.push(`i2="${rs.i2}"`);
        if (rs.i3 !== undefined) rsAttrs.push(`i3="${rs.i3}"`);
        if (rs.i4 !== undefined) rsAttrs.push(`i4="${rs.i4}"`);
        if (rs.ref) rsAttrs.push(`ref="${escapeXml(rs.ref)}"`);
        if (rs.name) rsAttrs.push(`name="${escapeXml(rs.name)}"`);
        if (rs.sheet) rsAttrs.push(`sheet="${escapeXml(rs.sheet)}"`);
        if (rs.rId) rsAttrs.push(`r:id="${escapeXml(rs.rId)}"`);
        conParts.push(`<rangeSet ${rsAttrs.join(" ")}/>`);
      }
      conParts.push("</rangeSets></consolidation></cacheSource>");
      p.push(conParts.join(""));
    } else {
      p.push(
        `<cacheSource type="worksheet">` +
          `<worksheetSource ref="${escapeXml(this.sourceRef)}" sheet="${escapeXml(this.sourceSheet)}"/>` +
          `</cacheSource>`,
      );
    }

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
        // String/categorical/mixed field: list unique values in sharedItems
        // Detect if field contains dates or nulls for correct metadata attributes
        let hasDate = false;
        let hasMissing = false;
        for (const v of uniqueVals) {
          if (v instanceof Date) hasDate = true;
          if (v === null) hasMissing = true;
        }
        const siAttrs: string[] = [`count="${uniqueVals.length}"`];
        if (hasDate) siAttrs.push('containsDate="1"');
        if (hasMissing) siAttrs.push('containsBlank="1"');

        p.push(
          `<cacheField name="${escapeXml(fieldName)}" numFmtId="0"><sharedItems ${siAttrs.join(" ")}>`,
        );
        for (const v of uniqueVals) {
          if (v === null) {
            p.push("<m/>");
          } else if (v instanceof Date) {
            p.push(`<d v="${v.toISOString().replace(/\.\d{3}Z$/, "Z")}"/>`);
          } else {
            p.push(`<s v="${escapeXml(String(v))}"/>`);
          }
        }
        p.push("</sharedItems>");

        // fieldGroup (CT_FieldGroup, after sharedItems per XSD sequence)
        const fg = this.cacheDefOpts?.fieldGroups?.get(i);
        if (fg) {
          const fgParts: string[] = ["<fieldGroup"];
          if (fg.parent !== undefined) fgParts.push(` par="${fg.parent}"`);
          if (fg.base !== undefined) fgParts.push(` base="${fg.base}"`);
          fgParts.push(">");
          // rangePr (CT_RangePr)
          if (fg.rangePr) {
            const rp = fg.rangePr;
            const rpAttrs: string[] = [];
            if (rp.autoStart === false) rpAttrs.push('autoStart="0"');
            if (rp.autoEnd === false) rpAttrs.push('autoEnd="0"');
            if (rp.groupBy && rp.groupBy !== "range") rpAttrs.push(`groupBy="${rp.groupBy}"`);
            if (rp.startNum !== undefined) rpAttrs.push(`startNum="${rp.startNum}"`);
            if (rp.endNum !== undefined) rpAttrs.push(`endNum="${rp.endNum}"`);
            if (rp.startDate) rpAttrs.push(`startDate="${escapeXml(rp.startDate)}"`);
            if (rp.endDate) rpAttrs.push(`endDate="${escapeXml(rp.endDate)}"`);
            if (rp.groupInterval !== undefined) rpAttrs.push(`groupInterval="${rp.groupInterval}"`);
            fgParts.push(`<rangePr${rpAttrs.length ? " " + rpAttrs.join(" ") : ""}/>`);
          }
          // discretePr (CT_DiscretePr)
          if (fg.discretePr && fg.discretePr.length > 0) {
            fgParts.push(`<discretePr count="${fg.discretePr.length}">`);
            for (const idx of fg.discretePr) {
              fgParts.push(`<x v="${idx}"/>`);
            }
            fgParts.push("</discretePr>");
          }
          // groupItems (CT_GroupItems)
          if (fg.groupItems && fg.groupItems.length > 0) {
            fgParts.push(`<groupItems count="${fg.groupItems.length}">`);
            for (const gi of fg.groupItems) {
              fgParts.push(`<s v="${escapeXml(gi)}"/>`);
            }
            fgParts.push("</groupItems>");
          }
          fgParts.push("</fieldGroup>");
          p.push(fgParts.join(""));
        }

        p.push("</cacheField>");
      }
    }

    p.push("</cacheFields>");

    // mpMap elements — rendered once outside cacheFields (CT_PivotCacheDefinition sequence)
    if (this.cacheDefOpts?.mpMaps) {
      for (const mp of this.cacheDefOpts.mpMaps) {
        p.push(`<mpMap x="${mp.x}"/>`);
      }
    }

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

        // groupLevels and fieldsUsage (optional children)
        const hasGroupLevels = ch.groupLevels && ch.groupLevels.length > 0;
        const hasFieldsUsage = ch.fieldsUsage && ch.fieldsUsage.length > 0;

        if (hasGroupLevels || hasFieldsUsage) {
          p.push(`<cacheHierarchy ${chAttrs.join(" ")}>`);
          // fieldsUsage (CT_FieldsUsage)
          if (hasFieldsUsage) {
            const fuParts = [`<fieldsUsage count="${ch.fieldsUsage!.length}">`];
            for (const fu of ch.fieldsUsage!) {
              fuParts.push(`<fieldUsage v="${fu.value}"/>`);
            }
            fuParts.push("</fieldsUsage>");
            p.push(fuParts.join(""));
          }
          // groupLevels (CT_GroupLevels)
          if (hasGroupLevels) {
            const glParts = [`<groupLevels count="${ch.groupLevels!.length}">`];
            for (const gl of ch.groupLevels!) {
              const glAttrs = [
                `uniqueName="${escapeXml(gl.uniqueName)}"`,
                `caption="${escapeXml(gl.caption)}"`,
              ];
              if (gl.user) glAttrs.push('user="1"');
              if (gl.customRollUp) glAttrs.push('customRollUp="1"');
              if (gl.groups && gl.groups.length > 0) {
                glParts.push(
                  `<groupLevel ${glAttrs.join(" ")}><groups count="${gl.groups.length}">`,
                );
                for (const lg of gl.groups) {
                  const lgAttrs = [
                    `name="${escapeXml(lg.name)}"`,
                    `uniqueName="${escapeXml(lg.uniqueName)}"`,
                    `caption="${escapeXml(lg.caption)}"`,
                  ];
                  if (lg.uniqueParent) lgAttrs.push(`uniqueParent="${escapeXml(lg.uniqueParent)}"`);
                  if (lg.id !== undefined) lgAttrs.push(`id="${lg.id}"`);
                  glParts.push(
                    `<group ${lgAttrs.join(" ")}><groupMembers count="${lg.members.length}">`,
                  );
                  for (const gm of lg.members) {
                    const gmAttrs = [`uniqueName="${escapeXml(gm.uniqueName)}"`];
                    if (gm.group) gmAttrs.push('group="1"');
                    glParts.push(`<groupMember ${gmAttrs.join(" ")}/>`);
                  }
                  glParts.push("</groupMembers></group>");
                }
                glParts.push("</groups></groupLevel>");
              } else {
                glParts.push(`<groupLevel ${glAttrs.join(" ")}/>`);
              }
            }
            glParts.push("</groupLevels>");
            p.push(glParts.join(""));
          }
          p.push("</cacheHierarchy>");
        } else {
          p.push(`<cacheHierarchy ${chAttrs.join(" ")}/>`);
        }
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

    // maps (CT_MeasureDimensionMaps, after measureGroups per XSD)
    if (
      this.cacheDefOpts?.measureDimensionMaps &&
      this.cacheDefOpts.measureDimensionMaps.length > 0
    ) {
      const mdm = this.cacheDefOpts.measureDimensionMaps;
      p.push(`<maps count="${mdm.length}">`);
      for (const m of mdm) {
        const mAttrs: string[] = [];
        if (m.measureGroup !== undefined) mAttrs.push(`measureGroup="${m.measureGroup}"`);
        if (m.dimension !== undefined) mAttrs.push(`dimension="${m.dimension}"`);
        p.push(`<map ${mAttrs.join(" ")}/>`);
      }
      p.push("</maps>");
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

    // tupleCache with entries, sets and serverFormats (optional)
    const cd = this.cacheDefOpts;
    const hasEntries = cd?.entries && cd.entries.length > 0;
    const hasSets = cd?.sets && cd.sets.length > 0;
    const hasServerFormats = cd?.serverFormats && cd.serverFormats.length > 0;
    const hasQueryCache = cd?.queryCache && cd.queryCache.length > 0;
    if (hasEntries || hasSets || hasServerFormats || hasQueryCache) {
      p.push("<tupleCache>");
      // entries (CT_PCDSDTCEntries)
      if (hasEntries) {
        const entParts: string[] = [`<entries count="${cd!.entries!.length}">`];
        for (const ent of cd!.entries!) {
          if (ent.type === "m") {
            entParts.push("<m/>");
          } else if (ent.type === "e" && ent.value !== undefined) {
            entParts.push(`<e v="${ent.value}"/>`);
          } else if (ent.value !== undefined) {
            entParts.push(`<${ent.type} v="${ent.value}"/>`);
          }
        }
        entParts.push("</entries>");
        p.push(entParts.join(""));
      }
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
      // queryCache (CT_QueryCache)
      if (hasQueryCache) {
        const qc = cd!.queryCache!;
        p.push(`<queryCache count="${qc.length}">`);
        for (const q of qc) {
          let qInner = "";
          if (q.tpls && q.tpls.length > 0) {
            qInner = `<tpls count="${q.tpls.length}">`;
            for (const tpl of q.tpls) {
              if (tpl.items && tpl.items.length > 0) {
                qInner += `<tpl>${tpl.items.map((i) => `<x v="${i}"/>`).join("")}</tpl>`;
              } else {
                qInner += "<tpl/>";
              }
            }
            qInner += "</tpls>";
          }
          if (qInner) {
            p.push(`<query mdx="${escapeXml(q.mdx)}">${qInner}</query>`);
          } else {
            p.push(`<query mdx="${escapeXml(q.mdx)}"/>`);
          }
        }
        p.push("</queryCache>");
      }

      p.push("</tupleCache>");
    }

    p.push("</pivotCacheDefinition>");

    return p.join("");
  }
}
