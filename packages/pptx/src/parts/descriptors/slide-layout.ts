/**
 * Slide Layout (p:sldLayout) descriptor for PPTX.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrNum, findChild } from "@office-open/xml";
import type { SlideLayoutType } from "@parts/slide-layout";
import type { MasterPlaceholderPosition } from "@parts/slide-master";

// ── Types ──

export interface SlideLayoutDescriptorOptions {
  /** Pre-built slide layout XML string. */
  layout: string;
  /** Layout type detected from display name. */
  type?: SlideLayoutType | string;
  /** Placeholder positions extracted from layout shapes. */
  placeholders?: Record<string, MasterPlaceholderPosition | false>;
}

// ── Display name → SlideLayoutType mapping ──

const NAME_TO_TYPE: Record<string, SlideLayoutType> = {
  "Title Slide": "title",
  "Title and Content": "obj",
  "Section Header": "secHead",
  "Two Content": "twoObj",
  Comparison: "twoTxTwoObj",
  "Title Only": "titleOnly",
  Blank: "blank",
  "Content with Caption": "objTx",
  "Picture with Caption": "picTx",
  "Vertical Text": "vertTx",
  "Vertical Title and Text": "vertTitleAndTx",
  "Title and Text": "tx",
};

/** Placeholder type → LayoutPlaceholderOptions key mapping. */
const PH_TYPE_TO_KEY: Record<string, string> = {
  title: "title",
  ctrTitle: "title",
  body: "body",
  sub: "subtitle",
  dt: "date",
  ftr: "footer",
  sldNum: "slideNumber",
};

// ── Descriptor ──

export const slideLayoutDesc: CustomDescriptor<SlideLayoutDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return opts.layout;
  },

  parse(el, _ctx) {
    const result: Record<string, unknown> = {};

    const cSld = findChild(el, "p:cSld");
    if (cSld) {
      const name = attr(cSld, "name");
      if (name) {
        const mapped = NAME_TO_TYPE[name];
        result.type = mapped ?? name;
      }

      // Extract placeholder positions from spTree shapes
      const spTree = findChild(cSld, "p:spTree");
      if (spTree) {
        const placeholders: Record<string, MasterPlaceholderPosition | false> = {};
        for (const child of spTree.elements ?? []) {
          if (child.name !== "p:sp") continue;
          const nvSpPr = findChild(child, "p:nvSpPr");
          const nvPr = nvSpPr ? findChild(nvSpPr, "p:nvPr") : undefined;
          const ph = nvPr ? findChild(nvPr, "p:ph") : undefined;
          if (!ph) continue;

          const phType = attr(ph, "type") as string | undefined;
          const key = phType ? PH_TYPE_TO_KEY[phType] : undefined;
          if (!key) continue;

          // Check if placeholder has "noRot" style (used to mark as hidden in some templates)
          const hasHiddenAttr = attr(ph, "sz") === "0";
          if (hasHiddenAttr) {
            placeholders[key] = false;
            continue;
          }

          // Extract position from spPr > a:xfrm
          const spPr = findChild(child, "p:spPr");
          const xfrm = spPr ? findChild(spPr, "a:xfrm") : undefined;
          if (!xfrm) continue;

          const off = findChild(xfrm, "a:off");
          const ext = findChild(xfrm, "a:ext");
          if (!off || !ext) continue;

          const x = attrNum(off, "x");
          const y = attrNum(off, "y");
          const width = attrNum(ext, "cx");
          const height = attrNum(ext, "cy");
          if (x === undefined || y === undefined || width === undefined || height === undefined)
            continue;

          placeholders[key] = { x, y, width, height };
        }
        if (Object.keys(placeholders).length > 0) result.placeholders = placeholders;
      }
    }

    return result as unknown as SlideLayoutDescriptorOptions;
  },
};
