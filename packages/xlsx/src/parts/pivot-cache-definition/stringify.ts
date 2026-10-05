import { escapeXml } from "@office-open/xml";

import {
  profilePivotFields,
  stringifyOlapPr,
  type PivotSourceData,
  type TupleOptions,
} from "../pivot/pivot-utils";
import { calculatedItemsXml, calculatedMembersXml } from "./calculated";
import type { PivotCacheDefinitionOptions } from "./types";

export function stringifyPivotCacheDefinition(
  sourceRef: string,
  sourceSheet: string,
  sourceData: PivotSourceData,
  recordsRid: string,
  cacheDefOpts?: PivotCacheDefinitionOptions,
): string {
  return stringifyPivotCacheDef(sourceRef, sourceSheet, sourceData, recordsRid, cacheDefOpts);
}

export function stringifyPivotCacheDefinitionOptions(options: PivotCacheDefinitionOptions): string {
  if (options.source) return options.source;
  const root = toPivotXmlElement(options, "pivotCacheDefinition");
  return stringifyElement(root);
}

function toPivotXmlElement(options: PivotCacheDefinitionOptions, name: string) {
  const attributes: Record<string, string | number | boolean | undefined> = {};
  const childKeys = new Set([
    "source",
    "recordsRelationshipSourceId",
    "recordsRelationshipTarget",
    "relationshipOrder",
    "externalRelationships",
    "cacheFields",
    "cacheSource",
    "cacheHierarchies",
    "kpis",
    "tupleCacheData",
    "calculatedItems",
    "calculatedMembers",
    "dimensions",
    "measureGroups",
    "maps",
    "extLst",
    "olap",
    "consolidation",
    "entries",
    "sets",
    "serverFormats",
    "queryCache",
    "fieldGroups",
    "mpMaps",
    "measureDimensionMaps",
    "cacheFieldOverrides",
  ]);
  for (const [key, value] of Object.entries(options as Record<string, unknown>)) {
    if (value === undefined || childKeys.has(key)) continue;
    if (key === "recordsRelationshipId") attributes["r:id"] = value as string;
    else attributes[key] = value as string | number | boolean | undefined;
  }
  return { name, attributes, children: [] };
}

function stringifyElement(element: ReturnType<typeof toPivotXmlElement>): string {
  const attrs = Object.entries(element.attributes)
    .filter(([, value]) => value !== undefined)
    .map(([key, value]) => `${key}="${escapeXml(String(value))}"`);
  if (!element.children.length)
    return `<${element.name}${attrs.length ? " " + attrs.join(" ") : ""}/>`;
  return `<${element.name}${attrs.length ? " " + attrs.join(" ") : ""}>${element.children.map(stringifyElement).join("")}</${element.name}>`;
}

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
    if (cd.refreshOnLoad) rootAttrs.push('refreshOnLoad="1"');
    if (cd.tupleCache) rootAttrs.push('tupleCache="1"');
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
      if (rs.pageFieldIndex1 !== undefined) rsAttrs.push(`i1="${rs.pageFieldIndex1}"`);
      if (rs.pageFieldIndex2 !== undefined) rsAttrs.push(`i2="${rs.pageFieldIndex2}"`);
      if (rs.pageFieldIndex3 !== undefined) rsAttrs.push(`i3="${rs.pageFieldIndex3}"`);
      if (rs.pageFieldIndex4 !== undefined) rsAttrs.push(`i4="${rs.pageFieldIndex4}"`);
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

  // cacheFields — one profile pass feeds every field's sharedItems
  const fieldNames = sourceData.fieldNames;
  const profiles = profilePivotFields(sourceData);
  p.push(`<cacheFields count="${fieldNames.length}">`);

  for (let i = 0; i < fieldNames.length; i++) {
    const fieldName = fieldNames[i] ?? "";
    const { numeric, unique: uniqueVals } = profiles[i]!;

    if (numeric) {
      let min = profiles[i]!.min,
        max = profiles[i]!.max;
      if (!isFinite(min)) {
        min = 0;
        max = 0;
      }
      const allInteger = profiles[i]!.allInteger;

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
        if (cfOverride.sqlType !== undefined) cfExtraAttrs.push(`sqlType="${cfOverride.sqlType}"`);
        if (cfOverride.containsMixedTypes) siExtraAttrs.push('containsMixedTypes="1"');
        if (cfOverride.containsNonDate) siExtraAttrs.push('containsNonDate="1"');
        if (cfOverride.longText) siExtraAttrs.push('longText="1"');
      }

      p.push(
        `<cacheField name="${escapeXml(fieldName)}" ${cfExtraAttrs.length ? cfExtraAttrs.join(" ") + " " : ""}numFmtId="0">` +
          `<sharedItems containsSemiMixedTypes="0" containsString="0"` +
          ` containsNumber="1" containsInteger="${allInteger ? "1" : "0"}"` +
          ` minValue="${min}" maxValue="${max}" count="${uniqueVals.length}"${siExtraAttrs.length ? " " + siExtraAttrs.join(" ") : ""}>`,
      );
      for (const v of uniqueVals) {
        if (v === null) p.push("<m/>");
        else if (v instanceof Date) p.push(`<d v="${v.toISOString().replace(/\.\d{3}Z$/, "Z")}"/>`);
        else p.push(`<n v="${v}"/>`);
      }
      p.push("</sharedItems></cacheField>");
    } else {
      let hasDate = false,
        hasMissing = false;
      for (const v of uniqueVals) {
        if (v instanceof Date) hasDate = true;
        if (v === null) hasMissing = true;
      }
      const siAttrs: string[] = [`count="${uniqueVals.length}"`];
      if (hasDate) {
        siAttrs.push('containsDate="1"');
        let minD: Date | undefined;
        let maxD: Date | undefined;
        for (const v of uniqueVals) {
          if (!(v instanceof Date)) continue;
          if (!minD || v < minD) minD = v;
          if (!maxD || v > maxD) maxD = v;
        }
        const iso = (d: Date) => d.toISOString().replace(/\.\d{3}Z$/, "Z");
        if (minD) siAttrs.push(`minDate="${iso(minD)}"`);
        if (maxD) siAttrs.push(`maxDate="${iso(maxD)}"`);
      }
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
        if (cfOverride.sqlType !== undefined) cfExtraAttrs.push(`sqlType="${cfOverride.sqlType}"`);
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
        if (fg.range) {
          const rp = fg.range;
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
        if (fg.discrete && fg.discrete.length > 0) {
          fgParts.push(`<discretePr count="${fg.discrete.length}">`);
          for (const idx of fg.discrete) fgParts.push(`<x v="${idx}"/>`);
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
  const olapPrXml = cacheDefOpts?.olap ? stringifyOlapPr(cacheDefOpts.olap) : "";
  if (olapPrXml) p.push(olapPrXml);

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
        const olapAttrs =
          (ent.captionIndex !== undefined ? ` cp="${ent.captionIndex}"` : "") +
          (ent.propertyIndex !== undefined ? ` in="${ent.propertyIndex}"` : "") +
          (ent.underline ? ' un="1"' : "");
        if (ent.type === "m") entParts.push(`<m${olapAttrs}/>`);
        else if (ent.type === "e")
          entParts.push(`<e v="${escapeXml(String(ent.value ?? ""))}"${olapAttrs}/>`);
        else if (ent.value !== undefined)
          entParts.push(`<${ent.type} v="${ent.value}"${olapAttrs}/>`);
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
        const tplsXml = (tpls: TupleOptions[] | undefined) =>
          tpls && tpls.length > 0
            ? `<tpls count="${tpls.length}">${tpls
                .map((tpl) =>
                  tpl.items && tpl.items.length > 0
                    ? `<tpl>${tpl.items.map((x) => `<x v="${x}"/>`).join("")}</tpl>`
                    : "<tpl/>",
                )
                .join("")}</tpls>`
            : "";
        const sortByXml =
          s.sortByTuple && s.sortByTuple.length > 0
            ? `<sortByTuple>${s.sortByTuple
                .map((tpl) =>
                  tpl.items && tpl.items.length > 0
                    ? `<tpl>${tpl.items.map((x) => `<x v="${x}"/>`).join("")}</tpl>`
                    : "<tpl/>",
                )
                .join("")}</sortByTuple>`
            : "";
        const inner = tplsXml(s.tpls) + sortByXml;
        if (inner) p.push(`<set ${sAttrs.join(" ")}>${inner}</set>`);
        else p.push(`<set ${sAttrs.join(" ")}/>`);
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

  if (cd?.calculatedItems?.length) p.push(calculatedItemsXml(cd.calculatedItems));
  if (cd?.calculatedMembers?.length) p.push(calculatedMembersXml(cd.calculatedMembers));

  p.push("</pivotCacheDefinition>");
  return p.join("");
}

// ── Stringify: pivotCacheRecords ──
