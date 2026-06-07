/**
 * Slide Layout (p:sldLayout) descriptor for PPTX.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";

// ── Types ──

export interface SlideLayoutDescriptorOptions {
  /** Pre-built slide layout XML string. */
  layout: string;
}

// ── Descriptor ──

export const slideLayoutDesc: CustomDescriptor<SlideLayoutDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return opts.layout;
  },

  parse(_el, _ctx) {
    return {} as Partial<SlideLayoutDescriptorOptions>;
  },
};
