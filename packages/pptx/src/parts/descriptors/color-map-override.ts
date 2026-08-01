/**
 * Color mapping override (p:clrMapOvr) descriptor for PPTX.
 *
 * CT_ColorMappingOverride — appears on slides, layouts, and masters. Two
 * forms: inherit the master mapping (a:masterClrMapping) or override the
 * 12 theme-color slots explicitly (a:overrideClrMapping bg1="dk1" ...).
 *
 * @module
 */
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { findChild } from "@office-open/xml";

export type ColorMapOverrideOptions =
  | { kind: "master" }
  | { kind: "override"; mapping: Record<string, string> };

export const colorMapOverrideDesc: CustomDescriptor<ColorMapOverrideOptions | undefined> = {
  kind: "custom",

  stringify(opts, _ctx) {
    if (opts?.kind === "override") {
      const attrs = Object.entries(opts.mapping)
        .map(([k, v]) => `${k}="${v}"`)
        .join(" ");
      return `<p:clrMapOvr><a:overrideClrMapping${attrs ? " " + attrs : ""}/></p:clrMapOvr>`;
    }
    // Explicit master or undefined default — both emit masterClrMapping.
    return "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>";
  },

  parse(el, _ctx) {
    if (findChild(el, "a:masterClrMapping")) return { kind: "master" };
    const override = findChild(el, "a:overrideClrMapping");
    if (override?.attributes) {
      const mapping: Record<string, string> = {};
      for (const [k, v] of Object.entries(override.attributes)) {
        if (typeof v === "string") mapping[k] = v;
      }
      if (Object.keys(mapping).length > 0) return { kind: "override", mapping };
    }
    return undefined;
  },
};
