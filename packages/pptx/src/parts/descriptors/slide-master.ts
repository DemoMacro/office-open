/**
 * Slide Master (p:sldMaster) descriptor for PPTX.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";

// ── Types ──

export interface SlideMasterDescriptorOptions {
  /** Pre-built slide master XML string. */
  master: string;
}

// ── Descriptor ──

export const slideMasterDesc: CustomDescriptor<SlideMasterDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return opts.master;
  },

  parse(_el, _ctx) {
    return {} as Partial<SlideMasterDescriptorOptions>;
  },
};
