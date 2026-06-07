/**
 * Slide Layout (p:sldLayout) descriptor for PPTX.
 *
 * @module
 */

import type { SlideLayout } from "@file/slide-layout";
import type { CustomDescriptor } from "@office-open/core/descriptor";

// ── Types ──

export interface SlideLayoutDescriptorOptions {
  /** Pre-built slide layout object from File. */
  layout: SlideLayout;
}

// ── Descriptor ──

export const slideLayoutDesc: CustomDescriptor<SlideLayoutDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return opts.layout.serialize();
  },

  parse(_el, _ctx) {
    return {} as Partial<SlideLayoutDescriptorOptions>;
  },
};
