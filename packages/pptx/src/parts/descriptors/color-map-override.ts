/**
 * Color mapping override (p:clrMapOvr) descriptor for PPTX.
 *
 * CT_ColorMappingOverride appears on child slides. It either inherits the
 * master mapping or contains a complete a:overrideClrMapping.
 *
 * @module
 */
import { parseColorMapping, stringifyColorMapping } from "@office-open/core";
import type { ColorMappingOptions } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { findChild } from "@office-open/xml";

export type ColorMappingOverrideOptions =
  | { kind: "master" }
  | { kind: "override"; colorMapping: Partial<ColorMappingOptions> };

export const colorMappingOverrideDesc: CustomDescriptor<ColorMappingOverrideOptions | undefined> = {
  kind: "custom",

  stringify(opts, _ctx) {
    if (opts?.kind === "override") {
      return `<p:clrMapOvr>${stringifyColorMapping(opts.colorMapping, "a:overrideClrMapping")}</p:clrMapOvr>`;
    }
    return "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>";
  },

  parse(el, _ctx) {
    if (findChild(el, "a:masterClrMapping")) return { kind: "master" };
    const colorMapping = parseColorMapping(findChild(el, "a:overrideClrMapping"));
    return colorMapping ? { kind: "override", colorMapping } : undefined;
  },
};
