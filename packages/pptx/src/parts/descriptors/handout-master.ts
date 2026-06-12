/**
 * Handout Master (p:handoutMaster) descriptor for PPTX.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, findChild } from "@office-open/xml";
import { buildHandoutMasterXml } from "@parts/handout-master";
import type { HandoutMasterOptions, ColorMapOptions } from "@parts/handout-master";

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
    const result: Record<string, unknown> = {};
    const options: Record<string, unknown> = {};

    // colorMap
    const clrMap = findChild(el, "p:clrMap");
    if (clrMap) {
      const colorMap: Record<string, string> = {};
      for (const key of COLOR_MAP_KEYS) {
        const v = attr(clrMap, key);
        if (v !== undefined) colorMap[key] = v;
      }
      if (Object.keys(colorMap).length > 0) options.colorMap = colorMap;
    }

    // headerFooter
    const hf = findChild(el, "p:hf");
    if (hf) {
      const headerFooter: Record<string, boolean> = {};
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

    if (Object.keys(options).length > 0) result.options = options;
    return result as unknown as HandoutMasterDescriptorOptions;
  },
};
