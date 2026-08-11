/**
 * PivotTable — parse helpers.
 *
 * @module
 */

import { attr, attrNum, findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type { PivotAreaOptions, PivotAreaReferenceOptions } from "../pivot/pivot-utils";

export function parsePivotArea(el: XmlElement): Partial<PivotAreaOptions> {
  const result: Partial<PivotAreaOptions> = {};
  const field = attrNum(el, "field");
  if (field !== undefined) result.field = field;
  const typeVal = attr(el, "type");
  if (typeVal) result.type = typeVal as PivotAreaOptions["type"];
  if (String(attr(el, "dataOnly")) === "0") result.dataOnly = false;
  if (String(attr(el, "labelOnly")) === "1") result.labelOnly = true;
  if (String(attr(el, "grandRow")) === "1") result.grandRow = true;
  if (String(attr(el, "grandCol")) === "1") result.grandCol = true;
  if (String(attr(el, "cacheIndex")) === "1") result.cacheIndex = true;
  if (String(attr(el, "outline")) === "0") result.outline = false;
  if (attr(el, "offset")) result.offset = attr(el, "offset");
  if (String(attr(el, "collapsedLevelsAreSubtotals")) === "1")
    result.collapsedLevelsAreSubtotals = true;
  const axisVal = attr(el, "axis");
  if (axisVal) result.axis = axisVal as PivotAreaOptions["axis"];
  const fp = attrNum(el, "fieldPosition");
  if (fp !== undefined) result.fieldPosition = fp;
  const refsEl = findChild(el, "references");
  if (refsEl) {
    const refs: Partial<PivotAreaReferenceOptions>[] = [];
    for (const rEl of refsEl.elements ?? []) {
      if (rEl.name !== "reference") continue;
      const ref: Partial<PivotAreaReferenceOptions> = {};
      const rField = attrNum(rEl, "field");
      if (rField !== undefined) ref.field = rField;
      const rCount = attrNum(rEl, "count");
      if (rCount !== undefined) ref.count = rCount;
      if (String(attr(rEl, "selected")) === "0") ref.selected = false;
      if (String(attr(rEl, "byPosition")) === "1") ref.byPosition = true;
      if (String(attr(rEl, "relative")) === "1") ref.relative = true;
      if (String(attr(rEl, "defaultSubtotal")) === "1") ref.defaultSubtotal = true;
      const xArr: number[] = [];
      for (const xEl of rEl.elements ?? []) {
        if (xEl.name === "x") {
          const v = attrNum(xEl, "v");
          if (v !== undefined) xArr.push(v);
        }
      }
      if (xArr.length > 0) ref.x = xArr;
      refs.push(ref);
    }
    result.references = refs;
  }
  return result;
}
