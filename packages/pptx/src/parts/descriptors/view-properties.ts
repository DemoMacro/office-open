/**
 * View Properties (p:viewPr) descriptor for PPTX.
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrNum, findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import {
  buildViewPropsXml,
  type NormalViewOptions,
  type SlideViewOptions,
  type ViewPropertiesOptions,
} from "@parts/view-properties";

// ── Descriptor ──

export const viewPropsDesc: CustomDescriptor<ViewPropertiesOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return buildViewPropsXml(opts);
  },

  parse(el, _ctx) {
    return parseViewProperties(el);
  },
};

// ── Parse ──

function parseViewProperties(el: XmlElement): ViewPropertiesOptions {
  const result: Partial<ViewPropertiesOptions> = {};

  if (el.attributes) {
    const a = el.attributes;
    if (a["lastView"]) {
      const reverseMap: Record<string, string> = {
        sldView: "slideView",
        sldMasterView: "slideMasterView",
        notesView: "notesView",
        handoutView: "handoutView",
        outlineView: "outlineView",
        sldSorterView: "slideSorterView",
      };
      const mapped = reverseMap[a["lastView"]];
      if (mapped) result.lastView = mapped as ViewPropertiesOptions["lastView"];
    }
    if (a["showComments"] !== undefined)
      result.showComments = parseOnOff(a["showComments"]) ?? true;
  }

  // p:normalViewPr
  const normalViewPr = findChild(el, "p:normalViewPr");
  if (normalViewPr) {
    const nv: Partial<NormalViewOptions> = {};
    if (String(normalViewPr.attributes?.["showOutlineIcons"]) === "0") nv.showOutlineIcons = false;
    if (parseOnOff(normalViewPr.attributes?.["snapVertSplitter"])) nv.snapVertSplitter = true;
    const vertBarState = attr(normalViewPr, "vertBarState");
    if (vertBarState) nv.vertBarState = vertBarState as NormalViewOptions["vertBarState"];
    const horzBarState = attr(normalViewPr, "horzBarState");
    if (horzBarState) nv.horzBarState = horzBarState as NormalViewOptions["horzBarState"];
    if (parseOnOff(normalViewPr.attributes?.["preferSingleView"])) nv.preferSingleView = true;
    if (Object.keys(nv).length > 0) result.normalView = nv as NormalViewOptions;
  }

  // p:slideViewPr > p:cSldViewPr (zoom, guides, slide-view toggles)
  const slideViewPr = findChild(el, "p:slideViewPr");
  if (slideViewPr) {
    const cSldViewPr = findChild(slideViewPr, "p:cSldViewPr");
    if (cSldViewPr) {
      const sv: Partial<SlideViewOptions> = {};
      if (String(cSldViewPr.attributes?.["snapToGrid"]) === "0") sv.snapToGrid = false;
      if (parseOnOff(cSldViewPr.attributes?.["snapToObjects"])) sv.snapToObjects = true;
      if (parseOnOff(cSldViewPr.attributes?.["showGuides"])) sv.showGuides = true;

      const cViewPr = findChild(cSldViewPr, "p:cViewPr");
      if (cViewPr) {
        const varScale = attr(cViewPr, "varScale");
        if (varScale !== undefined) sv.varScale = parseOnOff(varScale) ?? true;
        const scale = findChild(cViewPr, "p:scale");
        const sx = scale ? findChild(scale, "a:sx") : undefined;
        if (sx) {
          const n = attrNum(sx, "n");
          const d = attrNum(sx, "d");
          if (n !== undefined) result.zoomScaleNumerator = n;
          if (d !== undefined) result.zoomScaleDenominator = d;
        }
      }

      if (Object.keys(sv).length > 0) result.slideView = sv as SlideViewOptions;

      const guideLst = findChild(cSldViewPr, "p:guideLst");
      if (guideLst) {
        const guides: ViewPropertiesOptions["guides"] = [];
        for (const guide of guideLst.elements ?? []) {
          if (guide.name !== "p:guide") continue;
          const entry: { orient?: "vert" | "horz"; pos?: number } = {};
          const orient = attr(guide, "orient");
          if (orient) entry.orient = orient as "vert" | "horz";
          const pos = attrNum(guide, "pos");
          if (pos !== undefined) entry.pos = pos;
          guides.push(entry);
        }
        if (guides.length > 0) result.guides = guides;
      }
    }
  }

  // p:outlineViewPr > p:sldLst
  const outlineViewPr = findChild(el, "p:outlineViewPr");
  if (outlineViewPr) {
    const sldLst = findChild(outlineViewPr, "p:sldLst");
    if (sldLst) {
      const slides: NonNullable<NonNullable<ViewPropertiesOptions["outlineView"]>["slides"]> = [];
      for (const sld of sldLst.elements ?? []) {
        if (sld.name !== "p:sld") continue;
        const entry: { rId: string; collapse?: boolean } = { rId: attr(sld, "r:id") ?? "" };
        if (parseOnOff(attr(sld, "collapse"))) entry.collapse = true;
        slides.push(entry);
      }
      if (slides.length > 0) result.outlineView = { slides };
    }
  }

  // p:sorterViewPr @showFormatting
  const sorterViewPr = findChild(el, "p:sorterViewPr");
  if (sorterViewPr && parseOnOff(sorterViewPr.attributes?.["showFormatting"])) {
    result.sorterView = { showFormatting: true };
  }

  // p:notesViewPr (presence flag — content is the default cSldViewPr)
  if (findChild(el, "p:notesViewPr")) result.notesView = true;

  const gridSpacing = findChild(el, "p:gridSpacing");
  if (gridSpacing?.attributes) {
    const cx = attrNum(gridSpacing, "cx");
    const cy = attrNum(gridSpacing, "cy");
    if (cx !== undefined || cy !== undefined) {
      result.gridSpacing = { cx: cx ?? 72008, cy: cy ?? 72008 };
    }
  }

  return result as ViewPropertiesOptions;
}
