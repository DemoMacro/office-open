/**
 * PivotTable — parse helpers.
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
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
  if (parseOnOff(attr(el, "labelOnly"))) result.labelOnly = true;
  if (parseOnOff(attr(el, "grandRow"))) result.grandRow = true;
  if (parseOnOff(attr(el, "grandCol"))) result.grandCol = true;
  if (parseOnOff(attr(el, "cacheIndex"))) result.cacheIndex = true;
  if (String(attr(el, "outline")) === "0") result.outline = false;
  if (attr(el, "offset")) result.offset = attr(el, "offset");
  if (parseOnOff(attr(el, "collapsedLevelsAreSubtotals")))
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
      if (parseOnOff(attr(rEl, "byPosition"))) ref.byPosition = true;
      if (parseOnOff(attr(rEl, "relative"))) ref.relative = true;
      if (parseOnOff(attr(rEl, "defaultSubtotal"))) ref.defaultSubtotal = true;
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
