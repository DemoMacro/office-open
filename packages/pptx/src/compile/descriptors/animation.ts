/**
 * Animation timing (p:timing) descriptor for PPTX.
 *
 * @module
 */

import type { AnimationEntry } from "@file/animation/timing";
import { SlideTiming } from "@file/animation/timing";
import type { CustomDescriptor } from "@office-open/core/descriptor";

import { parseTiming } from "../../parse/animation";

// ── Types ──

export interface TimingDescriptorOptions {
  entries: AnimationEntry[];
}

// ── Descriptor ──

export const timingDesc: CustomDescriptor<TimingDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    if (opts.entries.length === 0) return "";
    const timing = new SlideTiming(opts.entries);
    return timing.serialize();
  },

  parse(el, _ctx) {
    const animMap = parseTiming(el);
    const entries: AnimationEntry[] = [];
    for (const [spid, options] of animMap) {
      entries.push({ spid, options });
    }
    return { entries } as Partial<TimingDescriptorOptions>;
  },
};
