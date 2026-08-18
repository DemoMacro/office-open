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
  type CommonViewPropertiesOptions,
  type NormalViewOptions,
  type NormalViewPortionOptions,
  type NotesViewOptions,
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

/** Parse p:cViewPr (CT_CommonViewProperties): scale + origin + varScale. */
function parseCommonViewProperties(el: XmlElement): CommonViewPropertiesOptions {
  const view: CommonViewPropertiesOptions = {};
  if (parseOnOff(attr(el, "varScale"))) view.variableScale = true;
  const scale = findChild(el, "p:scale");
  if (scale) {
    const sx = findChild(scale, "a:sx");
    const n = sx ? attrNum(sx, "n") : undefined;
    const d = sx ? attrNum(sx, "d") : undefined;
    if (n !== undefined && d !== undefined) view.scale = { numerator: n, denominator: d };
  }
  const origin = findChild(el, "p:origin");
  if (origin) {
    const x = attrNum(origin, "x");
    const y = attrNum(origin, "y");
    if (x !== undefined || y !== undefined) view.origin = { x: x ?? 0, y: y ?? 0 };
  }
  return view;
}

/** Parse p:restoredLeft / p:restoredTop (CT_NormalViewPortion). */
function parseNormalViewPortion(el: XmlElement): NormalViewPortionOptions {
  const portion: NormalViewPortionOptions = { size: attrNum(el, "sz") ?? 0 };
  const autoAdjust = attrNum(el, "autoAdjust");
  if (autoAdjust !== undefined) portion.autoAdjust = autoAdjust === 1;
  return portion;
}

/** Read a p:guideLst's entries (shared by slide-view and notes-view guides). */
function parseGuideLst(guideLst: XmlElement): { orient?: "vert" | "horz"; pos?: number }[] {
  const guides: { orient?: "vert" | "horz"; pos?: number }[] = [];
  for (const guide of guideLst.elements ?? []) {
    if (guide.name !== "p:guide") continue;
    const entry: { orient?: "vert" | "horz"; pos?: number } = {};
    const orient = attr(guide, "orient");
    if (orient) entry.orient = orient as "vert" | "horz";
    const pos = attrNum(guide, "pos");
    if (pos !== undefined) entry.pos = pos;
    guides.push(entry);
  }
  return guides;
}

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
    const restoredLeft = findChild(normalViewPr, "p:restoredLeft");
    if (restoredLeft) nv.restoredLeft = parseNormalViewPortion(restoredLeft);
    const restoredTop = findChild(normalViewPr, "p:restoredTop");
    if (restoredTop) nv.restoredTop = parseNormalViewPortion(restoredTop);
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
      if (cViewPr) sv.view = parseCommonViewProperties(cViewPr);

      if (Object.keys(sv).length > 0) result.slideView = sv as SlideViewOptions;

      const guideLst = findChild(cSldViewPr, "p:guideLst");
      if (guideLst) {
        const guides = parseGuideLst(guideLst);
        if (guides.length > 0) result.guides = guides;
      }
    }
  }

  // p:outlineViewPr — cViewPr (required) + sldLst
  const outlineViewPr = findChild(el, "p:outlineViewPr");
  if (outlineViewPr) {
    const cViewPr = findChild(outlineViewPr, "p:cViewPr");
    if (cViewPr) {
      const view = parseCommonViewProperties(cViewPr);
      const sldLst = findChild(outlineViewPr, "p:sldLst");
      const slides: NonNullable<NonNullable<ViewPropertiesOptions["outlineView"]>["slides"]> = [];
      if (sldLst) {
        for (const sld of sldLst.elements ?? []) {
          if (sld.name !== "p:sld") continue;
          const entry: { rId: string; collapse?: boolean } = { rId: attr(sld, "r:id") ?? "" };
          if (parseOnOff(attr(sld, "collapse"))) entry.collapse = true;
          slides.push(entry);
        }
      }
      result.outlineView = slides.length > 0 ? { view, slides } : { view };
    }
  }

  // p:notesTextViewPr — cViewPr only
  const notesTextViewPr = findChild(el, "p:notesTextViewPr");
  if (notesTextViewPr) {
    const cViewPr = findChild(notesTextViewPr, "p:cViewPr");
    if (cViewPr) result.notesTextView = parseCommonViewProperties(cViewPr);
  }

  // p:notesViewPr — cSldViewPr (view + guides)
  const notesViewPr = findChild(el, "p:notesViewPr");
  if (notesViewPr) {
    const cSld = findChild(notesViewPr, "p:cSldViewPr");
    const nv: NotesViewOptions = {};
    const cViewPr = cSld ? findChild(cSld, "p:cViewPr") : undefined;
    if (cViewPr) nv.view = parseCommonViewProperties(cViewPr);
    const guideLst = cSld ? findChild(cSld, "p:guideLst") : undefined;
    if (guideLst) {
      const guides = parseGuideLst(guideLst);
      if (guides.length > 0) nv.guides = guides;
    }
    result.notesView = Object.keys(nv).length > 0 ? nv : true;
  }

  // p:sorterViewPr — cViewPr + @showFormatting
  const sorterViewPr = findChild(el, "p:sorterViewPr");
  if (sorterViewPr) {
    const cViewPr = findChild(sorterViewPr, "p:cViewPr");
    if (cViewPr) {
      const sorter: NonNullable<ViewPropertiesOptions["sorterView"]> = {
        view: parseCommonViewProperties(cViewPr),
      };
      if (String(sorterViewPr.attributes?.["showFormatting"]) === "0") {
        sorter.showFormatting = false;
      }
      result.sorterView = sorter;
    }
  }

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
