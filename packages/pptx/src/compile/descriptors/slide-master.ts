/**
 * Slide Master (p:sldMaster) descriptor for PPTX.
 *
 * @module
 */

import type { DefaultSlideMaster } from "@file/slide-master";
import type { CustomDescriptor } from "@office-open/core/descriptor";

// ── Types ──

export interface SlideMasterDescriptorOptions {
  /** Pre-built slide master object from File. */
  master: DefaultSlideMaster;
}

// ── Descriptor ──

export const slideMasterDesc: CustomDescriptor<SlideMasterDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return opts.master.serialize();
  },

  parse(_el, _ctx) {
    return {} as Partial<SlideMasterDescriptorOptions>;
  },
};
