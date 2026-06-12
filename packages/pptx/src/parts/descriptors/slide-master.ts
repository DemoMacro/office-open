/**
 * Slide Master (p:sldMaster) descriptor for PPTX.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { stringify as xmlStringify } from "@office-open/xml";

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

  parse(el, _ctx) {
    const result: Record<string, unknown> = {};
    result.master = xmlStringify(el);
    return result as unknown as SlideMasterDescriptorOptions;
  },
};
