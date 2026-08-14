/**
 * AutoFilter shared stringify/parse helpers.
 *
 * CT_AutoFilter is the same content model in both worksheet (CT_Worksheet) and
 * table (CT_Table), so the XML round-trip lives here once and both consumers
 * delegate to it.
 *
 * Reference: OOXML transitional, sml.xsd, CT_AutoFilter / CT_FilterColumn /
 * CT_Filters / CT_SortState.
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import { attrs, attr, attrNum, escapeXml, findChild, selfCloseElement } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type {
  AutoFilterOptions,
  ColorFilterOptions,
  CustomFilterOptions,
  DateGroupFilterOptions,
  DynamicFilterOptions,
  FilterItemsOptions,
  IconFilterOptions,
  SortCondition,
  SortStateOptions,
  Top10FilterOptions,
} from "./worksheet";

/**
 * Stringify an autoFilter value (shorthand ref string or structured options)
 * into its `<autoFilter>` element. Returns "" for undefined so callers can push
 * unconditionally.
 */
export function stringifyAutoFilter(af: string | AutoFilterOptions | undefined): string {
  if (af === undefined) return "";
  if (typeof af === "string") {
    return selfCloseElement("autoFilter", attrs({ ref: af }));
  }
  const inner: string[] = [];
  for (const t10 of af.top10 ?? []) {
    const fcAttrs: Record<string, string | number | boolean | undefined> = {
      colId: t10.colId,
    };
    if (t10.hiddenButton) fcAttrs.hiddenButton = 1;
    if (t10.showButton === false) fcAttrs.showButton = 0;
    const t10Attrs: Record<string, string | number | boolean | undefined> = { val: t10.val };
    if (t10.top === false) t10Attrs.top = 0;
    if (t10.percent) t10Attrs.percent = 1;
    if (t10.filterVal !== undefined) t10Attrs.filterVal = t10.filterVal;
    inner.push(`<filterColumn${attrs(fcAttrs)}><top10${attrs(t10Attrs)}/></filterColumn>`);
  }
  for (const cf of af.customFilters ?? []) {
    const fcAttrs: Record<string, string | number | boolean | undefined> = {
      colId: cf.colId,
    };
    if (cf.hiddenButton) fcAttrs.hiddenButton = 1;
    if (cf.showButton === false) fcAttrs.showButton = 0;
    const cfAttrs: Record<string, string | number | boolean | undefined> = {};
    if (cf.and) cfAttrs.and = 1;
    const filters: string[] = [];
    if (cf.val !== undefined) {
      const fAttrs: Record<string, string | number | boolean | undefined> = { val: cf.val };
      if (cf.operator) fAttrs.operator = cf.operator;
      filters.push(selfCloseElement("customFilter", attrs(fAttrs)));
    }
    if (cf.val2 !== undefined) {
      filters.push(selfCloseElement("customFilter", attrs({ val: cf.val2 })));
    }
    if (filters.length > 0) {
      inner.push(
        `<filterColumn${attrs(fcAttrs)}><customFilters${attrs(cfAttrs)}>${filters.join("")}</customFilters></filterColumn>`,
      );
    }
  }
  // Simple filters (CT_Filters)
  for (const fi of af.filters ?? []) {
    const fcAttrs: Record<string, string | number | boolean | undefined> = {
      colId: fi.colId,
    };
    const filtersAttrs: Record<string, string | number | boolean | undefined> = {};
    if (fi.blank) filtersAttrs.blank = 1;
    if (fi.calendarType) filtersAttrs.calendarType = fi.calendarType;
    const valParts = (fi.values ?? []).map((v) => `<filter val="${escapeXml(v)}"/>`);
    inner.push(
      `<filterColumn${attrs(fcAttrs)}><filters${attrs(filtersAttrs)}>${valParts.join("")}</filters></filterColumn>`,
    );
  }
  if (af.sort && af.sort.length > 0) {
    const sortParts: string[] = [];
    for (const sc of af.sort) {
      const scAttrs: Record<string, string | number | boolean | undefined> = { ref: sc.ref };
      if (sc.descending) scAttrs.descending = 1;
      if (sc.sortBy) scAttrs.sortBy = sc.sortBy;
      if (sc.customList) scAttrs.customList = sc.customList;
      if (sc.iconId !== undefined) scAttrs.iconId = sc.iconId;
      sortParts.push(selfCloseElement("sortCondition", attrs(scAttrs)));
    }
    const ssAttrs: Record<string, string | number | boolean | undefined> = { ref: af.ref };
    if (af.sortState?.columnSort) ssAttrs.columnSort = 1;
    if (af.sortState?.caseSensitive) ssAttrs.caseSensitive = 1;
    if (af.sortState?.sortMethod) ssAttrs.sortMethod = af.sortState.sortMethod;
    inner.push(`<sortState${attrs(ssAttrs)}>${sortParts.join("")}</sortState>`);
  }
  // Color filters
  for (const cf of af.colorFilters ?? []) {
    const cfAttrs: Record<string, string | number | boolean | undefined> = {};
    if (cf.dxfId !== undefined) cfAttrs.dxfId = cf.dxfId;
    if (cf.cellColor === false) cfAttrs.cellColor = 0;
    inner.push(`<filterColumn colId="${cf.colId}"><colorFilter${attrs(cfAttrs)}/></filterColumn>`);
  }
  // Icon filters
  for (const if_ of af.iconFilters ?? []) {
    const ifAttrs: Record<string, string | number | boolean | undefined> = {
      iconSet: if_.iconSet,
    };
    if (if_.iconId !== undefined) ifAttrs.iconId = if_.iconId;
    inner.push(`<filterColumn colId="${if_.colId}"><iconFilter${attrs(ifAttrs)}/></filterColumn>`);
  }
  // Dynamic filters
  for (const df of af.dynamicFilters ?? []) {
    const dfAttrs: Record<string, string | number | boolean | undefined> = { type: df.type };
    if (df.val !== undefined) dfAttrs.val = df.val;
    if (df.maxVal !== undefined) dfAttrs.maxVal = df.maxVal;
    if (df.valIso !== undefined) dfAttrs.valIso = df.valIso;
    if (df.maxValIso !== undefined) dfAttrs.maxValIso = df.maxValIso;
    inner.push(
      `<filterColumn colId="${df.colId}"><dynamicFilter${attrs(dfAttrs)}/></filterColumn>`,
    );
  }
  // Date group filters (CT_Filters owns dateGroupItem)
  for (const dg of af.dateGroupItems ?? []) {
    const dgAttrs: Record<string, string | number | boolean | undefined> = {
      dateTimeGrouping: dg.dateTimeGrouping,
    };
    if (dg.year !== undefined) dgAttrs.year = dg.year;
    if (dg.month !== undefined) dgAttrs.month = dg.month;
    if (dg.day !== undefined) dgAttrs.day = dg.day;
    if (dg.hour !== undefined) dgAttrs.hour = dg.hour;
    if (dg.minute !== undefined) dgAttrs.minute = dg.minute;
    if (dg.second !== undefined) dgAttrs.second = dg.second;
    inner.push(
      `<filterColumn colId="${dg.colId}"><filters><dateGroupItem${attrs(dgAttrs)}/></filters></filterColumn>`,
    );
  }
  if (inner.length > 0) {
    return `<autoFilter ref="${af.ref}">${inner.join("")}</autoFilter>`;
  }
  return selfCloseElement("autoFilter", attrs({ ref: af.ref }));
}

/**
 * Parse an `<autoFilter>` element. Ref-only filters round-trip as the shorthand
 * string (matches the generate API); structured object only when filter columns
 * or sort state exist. nativeTypeAttributes (xlsx parse path) coerces "1"/"0"
 * to numbers, so boolean checks use String() coercion.
 */
export function parseAutoFilter(afEl: XmlElement): string | AutoFilterOptions {
  const af: AutoFilterOptions = { ref: attr(afEl, "ref") ?? "" };
  for (const child of afEl.elements ?? []) {
    if (child.name === "filterColumn") {
      const colId = attrNum(child, "colId") ?? 0;
      const hiddenButton = parseOnOff(attr(child, "hiddenButton"));
      const showButtonOff = String(attr(child, "showButton")) === "0";
      for (const fc of child.elements ?? []) {
        if (fc.name === "top10") {
          af.top10 ??= [];
          const t: Top10FilterOptions = { colId, val: attrNum(fc, "val") ?? 0 };
          if (String(attr(fc, "top")) === "0") t.top = false;
          if (parseOnOff(attr(fc, "percent"))) t.percent = true;
          const fv = attrNum(fc, "filterVal");
          if (fv !== undefined) t.filterVal = fv;
          if (hiddenButton) t.hiddenButton = true;
          if (showButtonOff) t.showButton = false;
          af.top10.push(t);
        } else if (fc.name === "customFilters") {
          af.customFilters ??= [];
          const cf: CustomFilterOptions = { colId };
          if (parseOnOff(attr(fc, "and"))) cf.and = true;
          const cfs = (fc.elements ?? []).filter((e) => e.name === "customFilter");
          const first = cfs[0];
          if (first) {
            const op = attr(first, "operator");
            if (op) cf.operator = op as CustomFilterOptions["operator"];
            const v = attr(first, "val");
            if (v !== undefined) cf.val = v;
          }
          const second = cfs[1];
          if (second) {
            const v2 = attr(second, "val");
            if (v2 !== undefined) cf.val2 = v2;
          }
          if (hiddenButton) cf.hiddenButton = true;
          if (showButtonOff) cf.showButton = false;
          af.customFilters.push(cf);
        } else if (fc.name === "filters") {
          // CT_Filters holds filter[] + dateGroupItem[]
          const fi: FilterItemsOptions = { colId };
          if (parseOnOff(attr(fc, "blank"))) fi.blank = true;
          const cal = attr(fc, "calendarType");
          if (cal) fi.calendarType = cal;
          const values: string[] = [];
          for (const f of fc.elements ?? []) {
            if (f.name === "filter") {
              const v = attr(f, "val");
              if (v !== undefined) values.push(v);
            } else if (f.name === "dateGroupItem") {
              af.dateGroupItems ??= [];
              const dg: DateGroupFilterOptions = {
                colId,
                dateTimeGrouping: attr(
                  f,
                  "dateTimeGrouping",
                ) as DateGroupFilterOptions["dateTimeGrouping"],
              };
              const y = attrNum(f, "year");
              if (y !== undefined) dg.year = y;
              const mo = attrNum(f, "month");
              if (mo !== undefined) dg.month = mo;
              const d = attrNum(f, "day");
              if (d !== undefined) dg.day = d;
              const h = attrNum(f, "hour");
              if (h !== undefined) dg.hour = h;
              const mi = attrNum(f, "minute");
              if (mi !== undefined) dg.minute = mi;
              const s = attrNum(f, "second");
              if (s !== undefined) dg.second = s;
              af.dateGroupItems.push(dg);
            }
          }
          if (values.length > 0) fi.values = values;
          af.filters ??= [];
          af.filters.push(fi);
        } else if (fc.name === "colorFilter") {
          af.colorFilters ??= [];
          const cf2: ColorFilterOptions = { colId };
          const dxfId = attrNum(fc, "dxfId");
          if (dxfId !== undefined) cf2.dxfId = dxfId;
          if (String(attr(fc, "cellColor")) === "0") cf2.cellColor = false;
          af.colorFilters.push(cf2);
        } else if (fc.name === "iconFilter") {
          af.iconFilters ??= [];
          const if_: IconFilterOptions = { colId, iconSet: attrNum(fc, "iconSet") ?? 0 };
          const iconId = attrNum(fc, "iconId");
          if (iconId !== undefined) if_.iconId = iconId;
          af.iconFilters.push(if_);
        } else if (fc.name === "dynamicFilter") {
          af.dynamicFilters ??= [];
          const df: DynamicFilterOptions = {
            colId,
            type: attr(fc, "type") as DynamicFilterOptions["type"],
          };
          const v = attrNum(fc, "val");
          if (v !== undefined) df.val = v;
          const mv = attrNum(fc, "maxVal");
          if (mv !== undefined) df.maxVal = mv;
          const vi = attr(fc, "valIso");
          if (vi) df.valIso = vi;
          const mvi = attr(fc, "maxValIso");
          if (mvi) df.maxValIso = mvi;
          af.dynamicFilters.push(df);
        }
      }
    } else if (child.name === "sortState") {
      const ss: SortStateOptions = {};
      if (parseOnOff(attr(child, "columnSort"))) ss.columnSort = true;
      if (parseOnOff(attr(child, "caseSensitive"))) ss.caseSensitive = true;
      const sm = attr(child, "sortMethod");
      if (sm) ss.sortMethod = sm as SortStateOptions["sortMethod"];
      af.sortState = ss;
      for (const sc of child.elements ?? []) {
        if (sc.name !== "sortCondition") continue;
        af.sort ??= [];
        const cond: SortCondition = { ref: attr(sc, "ref") ?? "" };
        if (parseOnOff(attr(sc, "descending"))) cond.descending = true;
        const sb = attr(sc, "sortBy");
        if (sb) cond.sortBy = sb as SortCondition["sortBy"];
        const cl = attr(sc, "customList");
        if (cl) cond.customList = cl;
        const ii = attrNum(sc, "iconId");
        if (ii !== undefined) cond.iconId = ii;
        af.sort.push(cond);
      }
    }
  }
  const hasFilters =
    af.top10 !== undefined ||
    af.customFilters !== undefined ||
    af.filters !== undefined ||
    af.colorFilters !== undefined ||
    af.iconFilters !== undefined ||
    af.dynamicFilters !== undefined ||
    af.dateGroupItems !== undefined ||
    af.sortState !== undefined ||
    af.sort !== undefined;
  return hasFilters ? af : af.ref;
}

/** Locate and parse an autoFilter child under a parent element. */
export function findAndParseAutoFilter(parent: XmlElement): string | AutoFilterOptions | undefined {
  const afEl = findChild(parent, "autoFilter");
  return afEl ? parseAutoFilter(afEl) : undefined;
}
