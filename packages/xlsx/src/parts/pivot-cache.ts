/**
 * PivotCache descriptors for XLSX — generates xl/pivotCache/pivotCacheDefinition{N}.xml
 * and xl/pivotCache/pivotCacheRecords{N}.xml.
 *
 * Direct stringify/parse — no intermediate class.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { escapeXml, findChild, attr, attrNum } from "@office-open/xml";

import type { PivotCacheDefinitionOptions } from "./pivot/pivot-utils";
import type { PivotSourceData } from "./pivot/pivot-utils";
import { collectUniqueValues, isNumericField } from "./pivot/pivot-utils";

// ── Types ──

export interface PivotCacheRecordsDescriptorOptions {
  sourceData: PivotSourceData;
  /** Parsed records (from parse path) */
  records?: Record<string, unknown>[][];
}

export interface PivotCacheDefDescriptorOptions {
  sourceRef: string;
  sourceSheet: string;
  sourceData: PivotSourceData;
  recordsRid: string;
  cacheDefOpts?: PivotCacheDefinitionOptions;
}

// ── Descriptors ──

export const pivotCacheDefDesc: CustomDescriptor<PivotCacheDefDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return stringifyPivotCacheDef(
      opts.sourceRef,
      opts.sourceSheet,
      opts.sourceData,
      opts.recordsRid,
      opts.cacheDefOpts,
    );
  },

  parse(el, _ctx) {
    const result: Record<string, unknown> = {};
    const csEl = findChild(el, "cacheSource");
    if (csEl) {
      result.sourceType = attr(csEl, "type");
      const wssEl = findChild(csEl, "worksheetSource");
      if (wssEl) {
        const wss: Record<string, unknown> = {};
        if (attr(wssEl, "ref")) wss.ref = attr(wssEl, "ref");
        if (attr(wssEl, "sheet")) wss.sheet = attr(wssEl, "sheet");
        result.worksheetSource = wss;
      }
    }
    const cfEl = findChild(el, "cacheFields");
    if (cfEl) {
      const fields: Record<string, unknown>[] = [];
      for (const fEl of cfEl.elements ?? []) {
        if (fEl.name !== "cacheField") continue;
        const field: Record<string, unknown> = {};
        if (attr(fEl, "name")) field.name = attr(fEl, "name");
        if (attrNum(fEl, "numFmtId") !== undefined) field.numFmtId = attrNum(fEl, "numFmtId");
        const siEl = findChild(fEl, "sharedItems");
        if (siEl) {
          const items: (string | number)[] = [];
          for (const siChild of siEl.elements ?? []) {
            const v = attr(siChild, "v");
            if (v !== undefined) items.push(isNaN(Number(v)) ? v : Number(v));
          }
          field.sharedItems = items;
        }
        fields.push(field);
      }
      result.cacheFields = fields;
    }
    return result as Record<string, unknown>;
  },
};

export const pivotCacheRecordsDesc: CustomDescriptor<PivotCacheRecordsDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return stringifyPivotCacheRecords(opts.sourceData);
  },

  parse(el, _ctx) {
    const records: Record<string, unknown>[][] = [];
    for (const rEl of el.elements ?? []) {
      if (rEl.name !== "r") continue;
      const record: Record<string, unknown>[] = [];
      for (const fEl of rEl.elements ?? []) {
        const entry: Record<string, unknown> = {};
        if (fEl.name === "x") {
          entry.type = "string";
          entry.v = attrNum(fEl, "v") ?? 0;
        } else if (fEl.name === "n") {
          entry.type = "number";
          const v = attr(fEl, "v");
          entry.v = v !== undefined ? Number(v) : 0;
        } else if (fEl.name === "d") {
          entry.type = "date";
          entry.v = attr(fEl, "v") ?? "";
        } else if (fEl.name === "m") {
          entry.type = "missing";
        }
        record.push(entry);
      }
      records.push(record);
    }
    return { records } as Record<string, unknown>;
  },
};

// ── Stringify: pivotCacheDefinition ──

function stringifyPivotCacheDef(
  sourceRef: string,
  sourceSheet: string,
  sourceData: PivotSourceData,
  recordsRid: string,
  cacheDefOpts?: PivotCacheDefinitionOptions,
): string {
  const p: string[] = [];
  const rootAttrs: string[] = [
    'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
    `r:id="${escapeXml(recordsRid)}"`,
    `recordCount="${sourceData.records.length}"`,
    'createdVersion="6"',
    'refreshedVersion="6"',
    'minRefreshableVersion="3"',
  ];

  if (cacheDefOpts) {
    const cd = cacheDefOpts;
    if (cd.invalid) rootAttrs.push('invalid="1"');
    if (cd.saveData === false) rootAttrs.push('saveData="0"');
    if (cd.optimizeMemory) rootAttrs.push('optimizeMemory="1"');
    if (cd.enableRefresh === false) rootAttrs.push('enableRefresh="0"');
    if (cd.refreshedBy) rootAttrs.push(`refreshedBy="${escapeXml(cd.refreshedBy)}"`);
    if (cd.refreshedDate !== undefined) rootAttrs.push(`refreshedDate="${cd.refreshedDate}"`);
    if (cd.refreshedDateIso) rootAttrs.push(`refreshedDateIso="${escapeXml(cd.refreshedDateIso)}"`);
    if (cd.backgroundQuery) rootAttrs.push('backgroundQuery="1"');
    if (cd.missingItemsLimit !== undefined)
      rootAttrs.push(`missingItemsLimit="${cd.missingItemsLimit}"`);
    if (cd.upgradeOnRefresh) rootAttrs.push('upgradeOnRefresh="1"');
    if (cd.supportSubquery) rootAttrs.push('supportSubquery="1"');
    if (cd.supportAdvancedDrill) rootAttrs.push('supportAdvancedDrill="1"');
  }

  p.push(`<pivotCacheDefinition ${rootAttrs.join(" ")}>`);

  // cacheSource
  if (cacheDefOpts?.consolidation) {
    const con = cacheDefOpts.consolidation;
    const conParts: string[] = ['<cacheSource type="consolidation"><consolidation'];
    if (con.autoPage === false) conParts.push(' autoPage="0"');
    conParts.push(">");
    if (con.pages && con.pages.length > 0) {
      conParts.push(`<pages count="${con.pages.length}">`);
      for (const pg of con.pages) {
        const pgItems = pg.items ?? [];
        conParts.push(`<page${pgItems.length ? ` count="${pgItems.length}"` : ""}>`);
        for (const pi of pgItems) conParts.push(`<pageItem name="${escapeXml(pi.name)}"/>`);
        conParts.push("</page>");
      }
      conParts.push("</pages>");
    }
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
        `<worksheetSource ref="${escapeXml(sourceRef)}" sheet="${escapeXml(sourceSheet)}"/>` +
        `</cacheSource>`,
    );
  }

  // cacheFields
  const fieldNames = sourceData.fieldNames;
  p.push(`<cacheFields count="${fieldNames.length}">`);

  for (let i = 0; i < fieldNames.length; i++) {
    const fieldName = fieldNames[i];
    const numeric = isNumericField(sourceData.records, i);
    const uniqueVals = collectUniqueValues(sourceData.records, i);

    if (numeric) {
      let min = Infinity,
        max = -Infinity;
      for (const row of sourceData.records) {
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
      const allInteger = sourceData.records.every(
        (row) => typeof row[i] === "number" && Number.isInteger(row[i]),
      );

      const cfOverride = cacheDefOpts?.cacheFieldOverrides?.get(i);
      const cfExtraAttrs: string[] = [];
      const siExtraAttrs: string[] = [];
      if (cfOverride) {
        if (cfOverride.databaseField) cfExtraAttrs.push('databaseField="1"');
        if (cfOverride.level !== undefined) cfExtraAttrs.push(`level="${cfOverride.level}"`);
        if (cfOverride.mappingCount !== undefined)
          cfExtraAttrs.push(`mappingCount="${cfOverride.mappingCount}"`);
        if (cfOverride.memberPropertyField !== undefined)
          cfExtraAttrs.push(`memberPropertyField="${cfOverride.memberPropertyField}"`);
        if (cfOverride.propertyName)
          cfExtraAttrs.push(`propertyName="${escapeXml(cfOverride.propertyName)}"`);
        if (cfOverride.serverField) cfExtraAttrs.push('serverField="1"');
        if (cfOverride.uniqueList) cfExtraAttrs.push('uniqueList="1"');
        if (cfOverride.containsMixedTypes) siExtraAttrs.push('containsMixedTypes="1"');
        if (cfOverride.containsNonDate) siExtraAttrs.push('containsNonDate="1"');
        if (cfOverride.longText) siExtraAttrs.push('longText="1"');
      }

      p.push(
        `<cacheField name="${escapeXml(fieldName)}" ${cfExtraAttrs.length ? cfExtraAttrs.join(" ") + " " : ""}numFmtId="0">` +
          `<sharedItems containsSemiMixedTypes="0" containsString="0"` +
          ` containsNumber="1" containsInteger="${allInteger ? "1" : "0"}"` +
          ` minValue="${min}" maxValue="${max}" count="${uniqueVals.length}"${siExtraAttrs.length ? " " + siExtraAttrs.join(" ") : ""}/>` +
          `</cacheField>`,
      );
    } else {
      let hasDate = false,
        hasMissing = false;
      for (const v of uniqueVals) {
        if (v instanceof Date) hasDate = true;
        if (v === null) hasMissing = true;
      }
      const siAttrs: string[] = [`count="${uniqueVals.length}"`];
      if (hasDate) siAttrs.push('containsDate="1"');
      if (hasMissing) siAttrs.push('containsBlank="1"');

      const cfOverride = cacheDefOpts?.cacheFieldOverrides?.get(i);
      const cfExtraAttrs: string[] = [];
      if (cfOverride) {
        if (cfOverride.databaseField) cfExtraAttrs.push('databaseField="1"');
        if (cfOverride.level !== undefined) cfExtraAttrs.push(`level="${cfOverride.level}"`);
        if (cfOverride.mappingCount !== undefined)
          cfExtraAttrs.push(`mappingCount="${cfOverride.mappingCount}"`);
        if (cfOverride.memberPropertyField !== undefined)
          cfExtraAttrs.push(`memberPropertyField="${cfOverride.memberPropertyField}"`);
        if (cfOverride.propertyName)
          cfExtraAttrs.push(`propertyName="${escapeXml(cfOverride.propertyName)}"`);
        if (cfOverride.serverField) cfExtraAttrs.push('serverField="1"');
        if (cfOverride.uniqueList) cfExtraAttrs.push('uniqueList="1"');
        if (cfOverride.containsMixedTypes) siAttrs.push('containsMixedTypes="1"');
        if (cfOverride.containsNonDate) siAttrs.push('containsNonDate="1"');
        if (cfOverride.longText) siAttrs.push('longText="1"');
      }

      p.push(
        `<cacheField name="${escapeXml(fieldName)}" ${cfExtraAttrs.length ? cfExtraAttrs.join(" ") + " " : ""}numFmtId="0"><sharedItems ${siAttrs.join(" ")}>`,
      );

      for (const v of uniqueVals) {
        if (v === null) p.push("<m/>");
        else if (v instanceof Date) p.push(`<d v="${v.toISOString().replace(/\.\d{3}Z$/, "Z")}"/>`);
        else p.push(`<s v="${escapeXml(String(v))}"/>`);
      }
      p.push("</sharedItems>");

      // fieldGroup
      const fg = cacheDefOpts?.fieldGroups?.get(i);
      if (fg) {
        const fgParts: string[] = ["<fieldGroup"];
        if (fg.parent !== undefined) fgParts.push(` par="${fg.parent}"`);
        if (fg.base !== undefined) fgParts.push(` base="${fg.base}"`);
        fgParts.push(">");
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
        if (fg.discretePr && fg.discretePr.length > 0) {
          fgParts.push(`<discretePr count="${fg.discretePr.length}">`);
          for (const idx of fg.discretePr) fgParts.push(`<x v="${idx}"/>`);
          fgParts.push("</discretePr>");
        }
        if (fg.groupItems && fg.groupItems.length > 0) {
          fgParts.push(`<groupItems count="${fg.groupItems.length}">`);
          for (const gi of fg.groupItems) fgParts.push(`<s v="${escapeXml(gi)}"/>`);
          fgParts.push("</groupItems>");
        }
        fgParts.push("</fieldGroup>");
        p.push(fgParts.join(""));
      }

      p.push("</cacheField>");
    }
  }
  p.push("</cacheFields>");

  // mpMap
  if (cacheDefOpts?.mpMaps) {
    for (const mp of cacheDefOpts.mpMaps) p.push(`<mpMap x="${mp.x}"/>`);
  }

  // olapPr
  if (cacheDefOpts?.olapPr) {
    const ol = cacheDefOpts.olapPr;
    const olAttrs: string[] = [];
    if (ol.local) olAttrs.push(` local="${escapeXml(ol.local)}"`);
    if (ol.localConnection) olAttrs.push(` localConnection="${escapeXml(ol.localConnection)}"`);
    if (ol.sendLocale) olAttrs.push(' sendLocale="1"');
    if (ol.rowDrillCount !== undefined) olAttrs.push(` rowDrillCount="${ol.rowDrillCount}"`);
    if (ol.colDrillCount !== undefined) olAttrs.push(` colDrillCount="${ol.colDrillCount}"`);
    if (ol.localRefresh) olAttrs.push(' localRefresh="1"');
    if (ol.serverFill === false) olAttrs.push(' serverFill="0"');
    if (ol.serverNumberFormat === false) olAttrs.push(' serverNumberFormat="0"');
    if (ol.serverFont === false) olAttrs.push(' serverFont="0"');
    if (ol.serverFontColor === false) olAttrs.push(' serverFontColor="0"');
    if (olAttrs.length > 0) p.push(`<olapPr${olAttrs.join("")}/>`);
  }

  // cacheHierarchies
  if (cacheDefOpts?.cacheHierarchies && cacheDefOpts.cacheHierarchies.length > 0) {
    const chs = cacheDefOpts.cacheHierarchies;
    p.push(`<cacheHierarchies count="${chs.length}">`);
    for (const ch of chs) {
      const chAttrs: string[] = [`uniqueName="${escapeXml(ch.uniqueName)}"`, `count="${ch.count}"`];
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
      if (ch.memberValueDatatype) chAttrs.push(`memberValueDatatype="${ch.memberValueDatatype}"`);
      if (ch.unbalanced) chAttrs.push('unbalanced="1"');
      if (ch.unbalancedGroup) chAttrs.push('unbalancedGroup="1"');

      const hasGL = ch.groupLevels && ch.groupLevels.length > 0;
      const hasFU = ch.fieldsUsage && ch.fieldsUsage.length > 0;
      if (hasGL || hasFU) {
        p.push(`<cacheHierarchy ${chAttrs.join(" ")}>`);
        if (hasFU) {
          const fuParts = [`<fieldsUsage count="${ch.fieldsUsage!.length}">`];
          for (const fu of ch.fieldsUsage!) fuParts.push(`<fieldUsage v="${fu.value}"/>`);
          fuParts.push("</fieldsUsage>");
          p.push(fuParts.join(""));
        }
        if (hasGL) {
          const glParts = [`<groupLevels count="${ch.groupLevels!.length}">`];
          for (const gl of ch.groupLevels!) {
            const glAttrs = [
              `uniqueName="${escapeXml(gl.uniqueName)}"`,
              `caption="${escapeXml(gl.caption)}"`,
            ];
            if (gl.user) glAttrs.push('user="1"');
            if (gl.customRollUp) glAttrs.push('customRollUp="1"');
            if (gl.groups && gl.groups.length > 0) {
              glParts.push(`<groupLevel ${glAttrs.join(" ")}><groups count="${gl.groups.length}">`);
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

  // kpis
  if (cacheDefOpts?.kpis && cacheDefOpts.kpis.length > 0) {
    p.push(`<kpis count="${cacheDefOpts.kpis.length}">`);
    for (const k of cacheDefOpts.kpis) {
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

  // measureGroups
  if (cacheDefOpts?.measureGroups && cacheDefOpts.measureGroups.length > 0) {
    p.push(`<measureGroups count="${cacheDefOpts.measureGroups.length}">`);
    for (const mg of cacheDefOpts.measureGroups) {
      p.push(`<measureGroup name="${escapeXml(mg.name)}" caption="${escapeXml(mg.caption)}"/>`);
    }
    p.push("</measureGroups>");
  }

  // maps
  if (cacheDefOpts?.measureDimensionMaps && cacheDefOpts.measureDimensionMaps.length > 0) {
    p.push(`<maps count="${cacheDefOpts.measureDimensionMaps.length}">`);
    for (const m of cacheDefOpts.measureDimensionMaps) {
      const mAttrs: string[] = [];
      if (m.measureGroup !== undefined) mAttrs.push(`measureGroup="${m.measureGroup}"`);
      if (m.dimension !== undefined) mAttrs.push(`dimension="${m.dimension}"`);
      p.push(`<map ${mAttrs.join(" ")}/>`);
    }
    p.push("</maps>");
  }

  // dimensions
  if (cacheDefOpts?.dimensions && cacheDefOpts.dimensions.length > 0) {
    p.push(`<dimensions count="${cacheDefOpts.dimensions.length}">`);
    for (const d of cacheDefOpts.dimensions) {
      const dAttrs: string[] = [
        `name="${escapeXml(d.name)}"`,
        `uniqueName="${escapeXml(d.uniqueName)}"`,
        `caption="${escapeXml(d.caption)}"`,
      ];
      if (d.measure) dAttrs.push('measure="1"');
      p.push(`<dimension ${dAttrs.join(" ")}/>`);
    }
    p.push("</dimensions>");
  }

  // tupleCache
  const cd = cacheDefOpts;
  const hasEntries = cd?.entries && cd.entries.length > 0;
  const hasSets = cd?.sets && cd.sets.length > 0;
  const hasSF = cd?.serverFormats && cd.serverFormats.length > 0;
  const hasQC = cd?.queryCache && cd.queryCache.length > 0;
  if (hasEntries || hasSets || hasSF || hasQC) {
    p.push("<tupleCache>");
    if (hasEntries) {
      const entParts: string[] = [`<entries count="${cd!.entries!.length}">`];
      for (const ent of cd!.entries!) {
        if (ent.type === "m") entParts.push("<m/>");
        else if (ent.value !== undefined) entParts.push(`<${ent.type} v="${ent.value}"/>`);
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
    if (hasSF) {
      p.push(`<serverFormats count="${cd!.serverFormats!.length}">`);
      for (const sf of cd!.serverFormats!) {
        const sfAttrs: string[] = [];
        if (sf.culture) sfAttrs.push(`culture="${escapeXml(sf.culture)}"`);
        if (sf.format) sfAttrs.push(`format="${escapeXml(sf.format)}"`);
        p.push(`<serverFormat ${sfAttrs.join(" ")}/>`);
      }
      p.push("</serverFormats>");
    }
    if (hasQC) {
      p.push(`<queryCache count="${cd!.queryCache!.length}">`);
      for (const q of cd!.queryCache!) {
        const qInner =
          q.tpls && q.tpls.length > 0
            ? `<tpls count="${q.tpls.length}">${q.tpls.map((tpl: { items?: number[] }) => (tpl.items && tpl.items.length > 0 ? `<tpl>${tpl.items.map((x: number) => `<x v="${x}"/>`).join("")}</tpl>` : "<tpl/>")).join("")}</tpls>`
            : "";
        if (qInner) p.push(`<query mdx="${escapeXml(q.mdx)}">${qInner}</query>`);
        else p.push(`<query mdx="${escapeXml(q.mdx)}"/>`);
      }
      p.push("</queryCache>");
    }
    p.push("</tupleCache>");
  }

  p.push("</pivotCacheDefinition>");
  return p.join("");
}

// ── Stringify: pivotCacheRecords ──

function stringifyPivotCacheRecords(sourceData: PivotSourceData): string {
  const numericFields = sourceData.fieldNames.map((_, i) => isNumericField(sourceData.records, i));
  const fieldIndexMaps: Map<string, number>[] = sourceData.fieldNames.map((_, i) => {
    if (numericFields[i]) return new Map<string, number>();
    const unique = collectUniqueValues(sourceData.records, i);
    const map = new Map<string, number>();
    for (let j = 0; j < unique.length; j++) map.set(String(unique[j]), j);
    return map;
  });

  const p: string[] = [];
  p.push(
    `<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sourceData.records.length}">`,
  );

  for (const row of sourceData.records) {
    p.push("<r>");
    for (let i = 0; i < row.length; i++) {
      const val = row[i];
      if (val === null) {
        p.push("<m/>");
      } else if (val instanceof Date) {
        p.push(`<d v="${val.toISOString().replace(/\.\d{3}Z$/, "Z")}"/>`);
      } else if (numericFields[i]) {
        p.push(`<n v="${val}"/>`);
      } else {
        p.push(`<x v="${fieldIndexMaps[i].get(String(val)) ?? 0}"/>`);
      }
    }
    p.push("</r>");
  }

  p.push("</pivotCacheRecords>");
  return p.join("");
}
