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
  CustomFiltersOptions,
  DateGroupFilterOptions,
  DynamicFilterOptions,
  FilterColumnOptions,
  FilterItemsOptions,
  IconFilterOptions,
  SortCondition,
  SortStateOptions,
  Top10FilterOptions,
} from "./worksheet";

function stringifyFilterColumn(fc: FilterColumnOptions): string {
  const fcAttrs: Record<string, string | number | boolean | undefined> = { colId: fc.colId };
  if (fc.hiddenButton) fcAttrs.hiddenButton = 1;
  if (fc.showButton === false) fcAttrs.showButton = 0;

  let inner: string;
  if (fc.top10) {
    const t10Attrs: Record<string, string | number | boolean | undefined> = { val: fc.top10.val };
    if (fc.top10.top === false) t10Attrs.top = 0;
    if (fc.top10.percent) t10Attrs.percent = 1;
    if (fc.top10.filterVal !== undefined) t10Attrs.filterVal = fc.top10.filterVal;
    inner = `<top10${attrs(t10Attrs)}/>`;
  } else if (fc.customFilters) {
    const cfAttrs: Record<string, string | number | boolean | undefined> = {};
    if (fc.customFilters.and) cfAttrs.and = 1;
    const entries = fc.customFilters.entries
      .map((e) => {
        const eAttrs: Record<string, string | number | boolean | undefined> = { val: e.val };
        if (e.operator) eAttrs.operator = e.operator;
        return selfCloseElement("customFilter", attrs(eAttrs));
      })
      .join("");
    inner = `<customFilters${attrs(cfAttrs)}>${entries}</customFilters>`;
  } else if (fc.filters) {
    const filtersAttrs: Record<string, string | number | boolean | undefined> = {};
    if (fc.filters.blank) filtersAttrs.blank = 1;
    if (fc.filters.calendarType) filtersAttrs.calendarType = fc.filters.calendarType;
    const valParts = (fc.filters.values ?? []).map((v) => `<filter val="${escapeXml(v)}"/>`);
    const dgParts = (fc.filters.dateGroupItems ?? []).map((dg) => {
      const dgAttrs: Record<string, string | number | boolean | undefined> = {
        dateTimeGrouping: dg.dateTimeGrouping,
      };
      if (dg.year !== undefined) dgAttrs.year = dg.year;
      if (dg.month !== undefined) dgAttrs.month = dg.month;
      if (dg.day !== undefined) dgAttrs.day = dg.day;
      if (dg.hour !== undefined) dgAttrs.hour = dg.hour;
      if (dg.minute !== undefined) dgAttrs.minute = dg.minute;
      if (dg.second !== undefined) dgAttrs.second = dg.second;
      return `<dateGroupItem${attrs(dgAttrs)}/>`;
    });
    inner = `<filters${attrs(filtersAttrs)}>${valParts.join("")}${dgParts.join("")}</filters>`;
  } else if (fc.colorFilter) {
    const cfAttrs: Record<string, string | number | boolean | undefined> = {};
    if (fc.colorFilter.dxfId !== undefined) cfAttrs.dxfId = fc.colorFilter.dxfId;
    if (fc.colorFilter.cellColor === false) cfAttrs.cellColor = 0;
    inner = `<colorFilter${attrs(cfAttrs)}/>`;
  } else if (fc.iconFilter) {
    const ifAttrs: Record<string, string | number | boolean | undefined> = {
      iconSet: fc.iconFilter.iconSet,
    };
    if (fc.iconFilter.iconId !== undefined) ifAttrs.iconId = fc.iconFilter.iconId;
    inner = `<iconFilter${attrs(ifAttrs)}/>`;
  } else if (fc.dynamicFilter) {
    const df = fc.dynamicFilter;
    const dfAttrs: Record<string, string | number | boolean | undefined> = { type: df.type };
    if (df.val !== undefined) dfAttrs.val = df.val;
    if (df.maxVal !== undefined) dfAttrs.maxVal = df.maxVal;
    if (df.valIso !== undefined) dfAttrs.valIso = df.valIso;
    if (df.maxValIso !== undefined) dfAttrs.maxValIso = df.maxValIso;
    inner = `<dynamicFilter${attrs(dfAttrs)}/>`;
  } else {
    inner = "";
  }

  if (inner) return `<filterColumn${attrs(fcAttrs)}>${inner}</filterColumn>`;
  return selfCloseElement("filterColumn", attrs(fcAttrs));
}

/**
 * Stringify a CT_SortState element. Shared by the autoFilter child and the
 * worksheet-level `sortState` sibling — one content model, one serializer.
 */
export function stringifySortStateXml(ss: SortStateOptions): string {
  const ssAttrs: Record<string, string | number | boolean | undefined> = { ref: ss.ref };
  if (ss.columnSort) ssAttrs.columnSort = 1;
  if (ss.caseSensitive) ssAttrs.caseSensitive = 1;
  if (ss.sortMethod && ss.sortMethod !== "none") ssAttrs.sortMethod = ss.sortMethod;
  const sortParts = ss.conditions.map((sc) => {
    const scAttrs: Record<string, string | number | boolean | undefined> = { ref: sc.ref };
    if (sc.descending) scAttrs.descending = 1;
    if (sc.sortBy && sc.sortBy !== "value") scAttrs.sortBy = sc.sortBy;
    if (sc.customList) scAttrs.customList = sc.customList;
    if (sc.dxfId !== undefined) scAttrs.dxfId = sc.dxfId;
    if (sc.iconSet && sc.iconSet !== "3Arrows") scAttrs.iconSet = sc.iconSet;
    if (sc.iconId !== undefined) scAttrs.iconId = sc.iconId;
    return selfCloseElement("sortCondition", attrs(scAttrs));
  });
  return `<sortState${attrs(ssAttrs)}>${sortParts.join("")}</sortState>`;
}

/** Parse a CT_SortState element. */
export function parseSortStateEl(el: XmlElement): SortStateOptions {
  const ss: SortStateOptions = { ref: attr(el, "ref") ?? "", conditions: [] };
  if (parseOnOff(attr(el, "columnSort"))) ss.columnSort = true;
  if (parseOnOff(attr(el, "caseSensitive"))) ss.caseSensitive = true;
  const sm = attr(el, "sortMethod");
  if (sm) ss.sortMethod = sm as SortStateOptions["sortMethod"];
  for (const sc of el.elements ?? []) {
    if (sc.name !== "sortCondition") continue;
    const cond: SortCondition = { ref: attr(sc, "ref") ?? "" };
    if (parseOnOff(attr(sc, "descending"))) cond.descending = true;
    const sb = attr(sc, "sortBy");
    if (sb) cond.sortBy = sb as SortCondition["sortBy"];
    const cl = attr(sc, "customList");
    if (cl) cond.customList = cl;
    const dxfId = attrNum(sc, "dxfId");
    if (dxfId !== undefined) cond.dxfId = dxfId;
    const is = attr(sc, "iconSet");
    if (is) cond.iconSet = is as SortCondition["iconSet"];
    const ii = attrNum(sc, "iconId");
    if (ii !== undefined) cond.iconId = ii;
    ss.conditions.push(cond);
  }
  return ss;
}

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
  for (const fc of af.columns ?? []) {
    inner.push(stringifyFilterColumn(fc));
  }
  if (af.sortState) {
    inner.push(stringifySortStateXml(af.sortState));
  }
  if (inner.length > 0) {
    return `<autoFilter ref="${escapeXml(af.ref)}">${inner.join("")}</autoFilter>`;
  }
  return selfCloseElement("autoFilter", attrs({ ref: af.ref }));
}

/** Parse a CT_Filters element (filter values + dateGroupItem children). */
function parseFiltersEl(fc: XmlElement): FilterItemsOptions {
  const fi: FilterItemsOptions = {};
  if (parseOnOff(attr(fc, "blank"))) fi.blank = true;
  const cal = attr(fc, "calendarType");
  if (cal) fi.calendarType = cal as FilterItemsOptions["calendarType"];
  const values: string[] = [];
  const dateGroupItems: DateGroupFilterOptions[] = [];
  for (const f of fc.elements ?? []) {
    if (f.name === "filter") {
      const v = attr(f, "val");
      if (v !== undefined) values.push(v);
    } else if (f.name === "dateGroupItem") {
      const dg: DateGroupFilterOptions = {
        dateTimeGrouping: attr(f, "dateTimeGrouping") as DateGroupFilterOptions["dateTimeGrouping"],
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
      dateGroupItems.push(dg);
    }
  }
  if (values.length > 0) fi.values = values;
  if (dateGroupItems.length > 0) fi.dateGroupItems = dateGroupItems;
  return fi;
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
      const fc: FilterColumnOptions = { colId: attrNum(child, "colId") ?? 0 };
      if (parseOnOff(attr(child, "hiddenButton"))) fc.hiddenButton = true;
      if (String(attr(child, "showButton")) === "0") fc.showButton = false;
      for (const el of child.elements ?? []) {
        if (el.name === "top10") {
          const t: Top10FilterOptions = { val: attrNum(el, "val") ?? 0 };
          if (String(attr(el, "top")) === "0") t.top = false;
          if (parseOnOff(attr(el, "percent"))) t.percent = true;
          const fv = attrNum(el, "filterVal");
          if (fv !== undefined) t.filterVal = fv;
          fc.top10 = t;
        } else if (el.name === "customFilters") {
          const cf: CustomFiltersOptions = { entries: [] };
          if (parseOnOff(attr(el, "and"))) cf.and = true;
          for (const entry of el.elements ?? []) {
            if (entry.name !== "customFilter") continue;
            const op = attr(entry, "operator");
            const v = attr(entry, "val");
            cf.entries.push({
              ...(op
                ? { operator: op as CustomFiltersOptions["entries"][number]["operator"] }
                : {}),
              val: v ?? "",
            });
          }
          fc.customFilters = cf;
        } else if (el.name === "filters") {
          fc.filters = parseFiltersEl(el);
        } else if (el.name === "colorFilter") {
          const colorFilter: ColorFilterOptions = {};
          const dxfId = attrNum(el, "dxfId");
          if (dxfId !== undefined) colorFilter.dxfId = dxfId;
          if (String(attr(el, "cellColor")) === "0") colorFilter.cellColor = false;
          fc.colorFilter = colorFilter;
        } else if (el.name === "iconFilter") {
          // @iconSet is ST_IconSetType (a string enum), not a number.
          const iconFilter: IconFilterOptions = {
            iconSet: (attr(el, "iconSet") ?? "3Arrows") as IconFilterOptions["iconSet"],
          };
          const iconId = attrNum(el, "iconId");
          if (iconId !== undefined) iconFilter.iconId = iconId;
          fc.iconFilter = iconFilter;
        } else if (el.name === "dynamicFilter") {
          const df: DynamicFilterOptions = {
            type: attr(el, "type") as DynamicFilterOptions["type"],
          };
          const v = attrNum(el, "val");
          if (v !== undefined) df.val = v;
          const mv = attrNum(el, "maxVal");
          if (mv !== undefined) df.maxVal = mv;
          const vi = attr(el, "valIso");
          if (vi) df.valIso = vi;
          const mvi = attr(el, "maxValIso");
          if (mvi) df.maxValIso = mvi;
          fc.dynamicFilter = df;
        }
      }
      (af.columns ??= []).push(fc);
    } else if (child.name === "sortState") {
      af.sortState = parseSortStateEl(child);
    }
  }
  const hasFilters = af.columns !== undefined || af.sortState !== undefined;
  return hasFilters ? af : af.ref;
}

/** Locate and parse an autoFilter child under a parent element. */
export function findAndParseAutoFilter(parent: XmlElement): string | AutoFilterOptions | undefined {
  const afEl = findChild(parent, "autoFilter");
  return afEl ? parseAutoFilter(afEl) : undefined;
}
