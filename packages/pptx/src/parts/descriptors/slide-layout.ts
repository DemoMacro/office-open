/**
 * Slide Layout (p:sldLayout) descriptor for PPTX.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, findChild } from "@office-open/xml";
import type { SlideLayoutType } from "@parts/slide-layout";

// ── Types ──

export interface SlideLayoutDescriptorOptions {
  /** Pre-built slide layout XML string. */
  layout: string;
  /** Layout type detected from display name. */
  type?: SlideLayoutType | string;
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
    }

    return result as Partial<SlideLayoutDescriptorOptions>;
  },
};
