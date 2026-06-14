/**
 * Handout Master (p:handoutMaster) descriptor for PPTX.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, findChild } from "@office-open/xml";
import { buildHandoutMasterXml } from "@parts/handout-master";
import type {
  ColorMapOptions,
  HandoutMasterOptions,
  HeaderFooterOptions,
} from "@parts/handout-master";

// ── Types ──

export interface HandoutMasterDescriptorOptions {
  options?: HandoutMasterOptions;
}

const COLOR_MAP_KEYS: (keyof ColorMapOptions)[] = [
  "bg1",
  "tx1",
  "bg2",
  "tx2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
];

// ── Descriptor ──

export const handoutMasterDesc: CustomDescriptor<HandoutMasterDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return buildHandoutMasterXml(opts.options);
  },

  parse(el, _ctx) {
    const options: HandoutMasterOptions = {};

    // colorMap
    const clrMap = findChild(el, "p:clrMap");
    if (clrMap) {
      const colorMap: ColorMapOptions = {};
      for (const key of COLOR_MAP_KEYS) {
        const v = attr(clrMap, key);
        if (v !== undefined) colorMap[key] = v;
      }
      if (Object.keys(colorMap).length > 0) options.colorMap = colorMap;
    }

    // headerFooter
    const hf = findChild(el, "p:hf");
    if (hf) {
      const headerFooter: HeaderFooterOptions = {};
      const dt = attr(hf, "dt");
      if (dt !== undefined) headerFooter.date = dt === "1";
      const hdr = attr(hf, "hdr");
      if (hdr !== undefined) headerFooter.header = hdr === "1";
      const ftr = attr(hf, "ftr");
      if (ftr !== undefined) headerFooter.footer = ftr === "1";
      const sldNum = attr(hf, "sldNum");
      if (sldNum !== undefined) headerFooter.slideNumber = sldNum === "1";
      if (Object.keys(headerFooter).length > 0) options.headerFooter = headerFooter;
    }

    return Object.keys(options).length > 0 ? { options } : {};
  },
};
