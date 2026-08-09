/**
 * Handout Master (p:handoutMaster) descriptor for PPTX.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { findChild } from "@office-open/xml";
import { buildHandoutMasterXml, parseColorMap, parseHeaderFooter } from "@parts/handout-master";
import type { HandoutMasterOptions } from "@parts/handout-master";

// ── Types ──

export interface HandoutMasterDescriptorOptions {
  options?: HandoutMasterOptions;
}

// ── Descriptor ──

export const handoutMasterDesc: CustomDescriptor<HandoutMasterDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return buildHandoutMasterXml(opts.options);
  },

  parse(el, _ctx) {
    const options: HandoutMasterOptions = {};

    const colorMap = parseColorMap(findChild(el, "p:clrMap"));
    if (colorMap) options.colorMap = colorMap;

    const headerFooter = parseHeaderFooter(findChild(el, "p:hf"));
    if (headerFooter) options.headerFooter = headerFooter;

    return Object.keys(options).length > 0 ? { options } : {};
  },
};
