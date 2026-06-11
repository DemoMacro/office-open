/**
 * Handout Master (p:handoutMaster) descriptor for PPTX.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { buildHandoutMasterXml } from "@parts/handout-master";
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

  parse(_el, _ctx) {
    return {} as Partial<HandoutMasterDescriptorOptions>;
  },
};
