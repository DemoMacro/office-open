/**
 * Color descriptors for DrawingML.
 *
 * @module
 */

import { escapeXml } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type { CustomDescriptor, ReadContext, WriteContext } from "../../descriptor";
import { stringify } from "../../descriptor";
import { emitAngle, emitPercent, parseAngle, parsePercent } from "../../util/converters";
import {
  ANGLE_TRANSFORMS,
  BOOLEAN_TRANSFORMS,
  createColorTransforms,
  PERCENT_TRANSFORMS,
} from "./color-transform";
import type { ColorTransformOptions } from "./color-transform";
import type { HslColorOptions } from "./hsl-color";
import type { PresetColorOptions } from "./preset-color";
import { PresetColor } from "./preset-color";
import type { RgbColorOptions } from "./rgb-color";
import type { ScRgbColorOptions } from "./sc-rgb-color";
import type { SchemeColorOptions } from "./scheme-color";
import { SchemeColor } from "./scheme-color";
import type { SolidFillOptions } from "./solid-fill";
import type { SystemColorOptions } from "./system-color";
import { SystemColor } from "./system-color";

// ── Color transform helpers ──

// The parse whitelist is the union of the three unit classes — the Sets in
// color-transform.ts are the single source of truth for the full key space.
const TRANSFORM_KEYS: readonly (keyof ColorTransformOptions & string)[] = [
  ...PERCENT_TRANSFORMS,
  ...ANGLE_TRANSFORMS,
  ...BOOLEAN_TRANSFORMS,
];

// Transform key classification (percent vs angle) lives in color-transform.ts
// so stringify and parse share one source of truth.

// Parse an ST_Percentage channel that may arrive as either a 1/1000th integer
// ("50000" = 50%) or a percent literal ("50%"); both are XSD-valid. Returns
// the caller-facing integer percent.
function parsePercentChannel(raw: string | number | undefined): number {
  if (raw === undefined) return 0;
  const s = typeof raw === "number" ? String(raw) : raw;
  if (s.endsWith("%")) return Number(s.slice(0, -1));
  return Number(s) / 1000;
}

function stringifyTransforms(opts: ColorTransformOptions): string {
  // Delegates to createColorTransforms so both paths emit in the caller's key
  // order (XSD transforms compose left to right — order is fidelity).
  return createColorTransforms(opts).join("");
}

function readTransforms(el: XmlElement): ColorTransformOptions | undefined {
  if (!el.elements || el.elements.length === 0) return undefined;
  const result: ColorTransformOptions = {};
  for (const child of el.elements) {
    if (!child.name || !child.name.startsWith("a:")) continue;
    const key = child.name.slice(2) as keyof ColorTransformOptions & string;
    if (!TRANSFORM_KEYS.includes(key)) continue;
    const val = child.attributes?.["val"];
    if (val !== undefined) {
      const raw = Number(val);
      // Percent → integer percent (÷1000); angle → degrees (÷60000); else raw.
      (result as Record<string, unknown>)[key] = PERCENT_TRANSFORMS.has(key)
        ? parsePercent(raw)
        : ANGLE_TRANSFORMS.has(key)
          ? parseAngle(raw)
          : raw;
    } else {
      (result as Record<string, unknown>)[key] = true;
    }
  }
  return Object.keys(result).length > 0 ? result : undefined;
}

// ── RgbColor descriptor ──

export const rgbColorDesc: CustomDescriptor<RgbColorOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const transforms = opts.transforms ? stringifyTransforms(opts.transforms) : "";
    if (transforms) {
      return `<a:srgbClr val="${escapeXml(opts.value)}">${transforms}</a:srgbClr>`;
    }
    return `<a:srgbClr val="${escapeXml(opts.value)}"/>`;
  },
  parse(el, _ctx) {
    const result: RgbColorOptions = { value: String(el.attributes?.["val"] ?? "") };
    const transforms = readTransforms(el);
    if (transforms) result.transforms = transforms;
    return result;
  },
};

// ── SchemeColor descriptor ──

export const schemeColorDesc: CustomDescriptor<SchemeColorOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const transforms = opts.transforms ? stringifyTransforms(opts.transforms) : "";
    if (transforms) {
      return `<a:schemeClr val="${escapeXml(opts.value)}">${transforms}</a:schemeClr>`;
    }
    return `<a:schemeClr val="${escapeXml(opts.value)}"/>`;
  },
  parse(el, _ctx) {
    const result: SchemeColorOptions = {
      value: String(el.attributes?.["val"] ?? "") as SchemeColorOptions["value"],
    };
    const transforms = readTransforms(el);
    if (transforms) result.transforms = transforms;
    return result;
  },
};

// ── HslColor descriptor ──

export const hslColorDesc: CustomDescriptor<HslColorOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const transforms = opts.transforms ? stringifyTransforms(opts.transforms) : "";
    const sat = emitPercent(opts.saturation);
    const lum = emitPercent(opts.luminance);
    const hue = emitAngle(opts.hue);
    if (transforms) {
      return `<a:hslClr hue="${hue}" sat="${sat}" lum="${lum}">${transforms}</a:hslClr>`;
    }
    return `<a:hslClr hue="${hue}" sat="${sat}" lum="${lum}"/>`;
  },
  parse(el, _ctx) {
    const result: HslColorOptions = {
      hue: parseAngle(Number(el.attributes?.["hue"] ?? 0)),
      saturation: parsePercent(Number(el.attributes?.["sat"] ?? 0)),
      luminance: parsePercent(Number(el.attributes?.["lum"] ?? 0)),
    };
    const transforms = readTransforms(el);
    if (transforms) result.transforms = transforms;
    return result;
  },
};

// ── SystemColor descriptor ──

export const systemColorDesc: CustomDescriptor<SystemColorOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const transforms = opts.transforms ? stringifyTransforms(opts.transforms) : "";
    const attrParts: string[] = [`val="${escapeXml(opts.value)}"`];
    if (opts.lastClr) attrParts.push(`lastClr="${escapeXml(opts.lastClr)}"`);
    const attrStr = attrParts.join(" ");
    if (transforms) {
      return `<a:sysClr ${attrStr}>${transforms}</a:sysClr>`;
    }
    return `<a:sysClr ${attrStr}/>`;
  },
  parse(el, _ctx) {
    const result: SystemColorOptions = {
      value: String(el.attributes?.["val"] ?? "") as SystemColorOptions["value"],
    };
    const lastClr = el.attributes?.["lastClr"];
    if (lastClr) result.lastClr = String(lastClr);
    const transforms = readTransforms(el);
    if (transforms) result.transforms = transforms;
    return result;
  },
};

// ── PresetColor descriptor ──

export const presetColorDesc: CustomDescriptor<PresetColorOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const transforms = opts.transforms ? stringifyTransforms(opts.transforms) : "";
    if (transforms) {
      return `<a:prstClr val="${escapeXml(opts.value)}">${transforms}</a:prstClr>`;
    }
    return `<a:prstClr val="${escapeXml(opts.value)}"/>`;
  },
  parse(el, _ctx) {
    const result: PresetColorOptions = {
      value: String(el.attributes?.["val"] ?? "") as PresetColorOptions["value"],
    };
    const transforms = readTransforms(el);
    if (transforms) result.transforms = transforms;
    return result;
  },
};

// ── ScRgbColor descriptor ──

export const scRgbColorDesc: CustomDescriptor<ScRgbColorOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const transforms = opts.transforms ? stringifyTransforms(opts.transforms) : "";
    const r = emitPercent(opts.r);
    const g = emitPercent(opts.g);
    const b = emitPercent(opts.b);
    if (transforms) {
      return `<a:scrgbClr r="${r}" g="${g}" b="${b}">${transforms}</a:scrgbClr>`;
    }
    return `<a:scrgbClr r="${r}" g="${g}" b="${b}"/>`;
  },
  parse(el, _ctx) {
    const result: ScRgbColorOptions = {
      r: parsePercentChannel(el.attributes?.["r"]),
      g: parsePercentChannel(el.attributes?.["g"]),
      b: parsePercentChannel(el.attributes?.["b"]),
    };
    const transforms = readTransforms(el);
    if (transforms) result.transforms = transforms;
    return result;
  },
};

// ── Color discrimination (SolidFillOptions) ──

const SYSTEM_COLOR_VALUES: ReadonlySet<string> = new Set(Object.values(SystemColor));
const PRESET_COLOR_VALUES: ReadonlySet<string> = new Set(Object.values(PresetColor));
const SCHEME_COLOR_VALUES: ReadonlySet<string> = new Set(Object.values(SchemeColor));

/** Stringify an EG_ColorChoice (direct color element, no `a:solidFill` wrapper).
 * Used for gradient stops, fg/bg clr, and effect colors. Replaces the former
 * `getColorDescriptor` which returned a polymorphic `CustomDescriptor<any>`. */
export function stringifyColorChoice(color: SolidFillOptions, ctx: WriteContext): string {
  if ("hue" in color) return stringify(hslColorDesc, color, ctx) ?? "";
  if ("r" in color) return stringify(scRgbColorDesc, color, ctx) ?? "";
  // Remaining variants (rgb/scheme/system/preset) share a `value` key.
  const colorValue = (color as { value: string }).value;
  if (SYSTEM_COLOR_VALUES.has(colorValue))
    return stringify(systemColorDesc, color as SystemColorOptions, ctx) ?? "";
  if (PRESET_COLOR_VALUES.has(colorValue))
    return stringify(presetColorDesc, color as PresetColorOptions, ctx) ?? "";
  if (SCHEME_COLOR_VALUES.has(colorValue))
    return stringify(schemeColorDesc, color as SchemeColorOptions, ctx) ?? "";
  return stringify(rgbColorDesc, color as RgbColorOptions, ctx) ?? "";
}

/**
 * Parse an EG_ColorChoice from an element's direct children. Handles all six
 * color element kinds (srgbClr/schemeClr/hslClr/sysClr/prstClr/scrgbClr) —
 * used both by {@link solidFillDesc} (under a:solidFill) and by fill
 * descriptors reading direct colors under a:gs / a:fgClr / a:bgClr.
 */
export function parseColorChoice(el: XmlElement, ctx: ReadContext): SolidFillOptions {
  if (!el.elements) return {} as SolidFillOptions;
  for (const child of el.elements) {
    switch (child.name) {
      case "a:srgbClr":
        return rgbColorDesc.parse(child, ctx);
      case "a:schemeClr":
        return schemeColorDesc.parse(child, ctx);
      case "a:hslClr":
        return hslColorDesc.parse(child, ctx);
      case "a:sysClr":
        return systemColorDesc.parse(child, ctx);
      case "a:prstClr":
        return presetColorDesc.parse(child, ctx);
      case "a:scrgbClr":
        return scRgbColorDesc.parse(child, ctx);
    }
  }
  return {} as SolidFillOptions;
}

// ── SolidFill descriptor ──

export const solidFillDesc: CustomDescriptor<SolidFillOptions> = {
  kind: "custom",
  stringify(color, ctx) {
    const inner = stringifyColorChoice(color, ctx);
    if (!inner) return undefined;
    return `<a:solidFill>${inner}</a:solidFill>`;
  },
  parse(el, _ctx) {
    return parseColorChoice(el, _ctx);
  },
};
