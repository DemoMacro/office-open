/**
 * Handout Master (p:handoutMaster) descriptor for PPTX.
 *
 * @module
 */

import { DefaultHandoutMaster } from "@file/handout-master";
import type { HandoutMasterOptions } from "@file/handout-master";
import type { CustomDescriptor } from "@office-open/core/descriptor";

// ── Types ──

export interface HandoutMasterDescriptorOptions {
  options?: HandoutMasterOptions;
}

// ── Descriptor ──

export const handoutMasterDesc: CustomDescriptor<HandoutMasterDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const master = new DefaultHandoutMaster(opts.options);
    return master.serialize();
  },

  parse(_el, _ctx) {
    return {} as Partial<HandoutMasterDescriptorOptions>;
  },
};
