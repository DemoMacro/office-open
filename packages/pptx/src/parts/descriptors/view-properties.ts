/**
 * View Properties (p:viewPr) descriptor for PPTX.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { findChild, attrNum } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { buildViewPropsXml, type ViewPropertiesOptions } from "@parts/view-properties";

// ── Types ──

export type ViewPropertiesDescriptorOptions = ViewPropertiesOptions;

// ── Descriptor ──

export const viewPropsDesc: CustomDescriptor<ViewPropertiesDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return buildViewPropsXml(opts);
  },

  parse(el, _ctx) {
    return parseViewProperties(el);
  },
};

// ── Parse ──

function parseViewProperties(el: XmlElement): ViewPropertiesDescriptorOptions {
  const result: Record<string, unknown> = {};

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
      if (mapped) result.lastView = mapped;
    }
    if (a["showComments"] !== undefined) result.showComments = a["showComments"] !== "0";
  }

  const gridSpacing = findChild(el, "p:gridSpacing");
  if (gridSpacing?.attributes) {
    const cx = attrNum(gridSpacing, "cx");
    const cy = attrNum(gridSpacing, "cy");
    if (cx !== undefined || cy !== undefined) {
      result.gridSpacing = { cx: cx ?? 72008, cy: cy ?? 72008 };
    }
  }

  return result as unknown as ViewPropertiesDescriptorOptions;
}
