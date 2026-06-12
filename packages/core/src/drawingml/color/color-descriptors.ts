/**
 * Color descriptors for DrawingML.
 *
 * @module
 */

import { escapeXml } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type { CustomDescriptor } from "../../descriptor";
import { stringify } from "../../descriptor";
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

const TRANSFORM_KEYS: readonly (keyof ColorTransformOptions & string)[] = [
  "tint",
  "shade",
  "comp",
  "inv",
  "gray",
  "alpha",
  "alphaOff",
  "alphaMod",
  "hue",
  "hueOff",
  "hueMod",
  "sat",
  "satOff",
  "satMod",
  "lum",
  "lumOff",
  "lumMod",
  "red",
  "redOff",
  "redMod",
  "green",
  "greenOff",
  "greenMod",
  "blue",
  "blueOff",
  "blueMod",
  "gamma",
  "invGamma",
];

function stringifyTransforms(opts: ColorTransformOptions): string {
  const parts: string[] = [];
  for (const key of TRANSFORM_KEYS) {
    const v = (opts as Record<string, unknown>)[key];
    if (v === undefined || v === false) continue;
    if (v === true) {
      parts.push(`<a:${key}/>`);
    } else {
      parts.push(`<a:${key} val="${escapeXml(String(v as number))}"/>`);
    }
  }
  return parts.join("");
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
      (result as Record<string, unknown>)[key] = Number(val);
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
    const result: Partial<RgbColorOptions> = {};
    result.value = String(el.attributes?.["val"] ?? "");
    const transforms = readTransforms(el);
    if (transforms) result.transforms = transforms;
    return result as RgbColorOptions;
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
    const result: Partial<SchemeColorOptions> = {};
    result.value = String(el.attributes?.["val"] ?? "") as SchemeColorOptions["value"];
    const transforms = readTransforms(el);
    if (transforms) result.transforms = transforms;
    return result as SchemeColorOptions;
  },
};

// ── HslColor descriptor ──

export const hslColorDesc: CustomDescriptor<HslColorOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const transforms = opts.transforms ? stringifyTransforms(opts.transforms) : "";
    if (transforms) {
      return `<a:hslClr hue="${opts.hue}" sat="${opts.saturation}" lum="${opts.luminance}">${transforms}</a:hslClr>`;
    }
    return `<a:hslClr hue="${opts.hue}" sat="${opts.saturation}" lum="${opts.luminance}"/>`;
  },
  parse(el, _ctx) {
    const result: Partial<HslColorOptions> = {};
    result.hue = Number(el.attributes?.["hue"] ?? 0);
    result.saturation = Number(el.attributes?.["sat"] ?? 0);
    result.luminance = Number(el.attributes?.["lum"] ?? 0);
    const transforms = readTransforms(el);
    if (transforms) result.transforms = transforms;
    return result as HslColorOptions;
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
    const result: Partial<SystemColorOptions> = {};
    result.value = String(el.attributes?.["val"] ?? "") as SystemColorOptions["value"];
    const lastClr = el.attributes?.["lastClr"];
    if (lastClr) result.lastClr = String(lastClr);
    const transforms = readTransforms(el);
    if (transforms) result.transforms = transforms;
    return result as SystemColorOptions;
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
    const result: Partial<PresetColorOptions> = {};
    result.value = String(el.attributes?.["val"] ?? "") as PresetColorOptions["value"];
    const transforms = readTransforms(el);
    if (transforms) result.transforms = transforms;
    return result as PresetColorOptions;
  },
};

// ── ScRgbColor descriptor ──

export const scRgbColorDesc: CustomDescriptor<ScRgbColorOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const transforms = opts.transforms ? stringifyTransforms(opts.transforms) : "";
    if (transforms) {
      return `<a:scrgbClr r="${escapeXml(opts.r)}" g="${escapeXml(opts.g)}" b="${escapeXml(opts.b)}">${transforms}</a:scrgbClr>`;
    }
    return `<a:scrgbClr r="${escapeXml(opts.r)}" g="${escapeXml(opts.g)}" b="${escapeXml(opts.b)}"/>`;
  },
  parse(el, _ctx) {
    const result: Partial<ScRgbColorOptions> = {};
    result.r = String(el.attributes?.["r"] ?? "");
    result.g = String(el.attributes?.["g"] ?? "");
    result.b = String(el.attributes?.["b"] ?? "");
    const transforms = readTransforms(el);
    if (transforms) result.transforms = transforms;
    return result as ScRgbColorOptions;
  },
};

// ── Color discrimination (SolidFillOptions) ──

const SYSTEM_COLOR_VALUES: ReadonlySet<string> = new Set(Object.values(SystemColor));
const PRESET_COLOR_VALUES: ReadonlySet<string> = new Set(Object.values(PresetColor));
const SCHEME_COLOR_VALUES: ReadonlySet<string> = new Set(Object.values(SchemeColor));

export function getColorDescriptor(color: SolidFillOptions): CustomDescriptor<any> {
  if ("hue" in color && "saturation" in color && "luminance" in color) return hslColorDesc;
  if ("r" in color && "g" in color && "b" in color) return scRgbColorDesc;
  const colorValue = (color as { value: string }).value;
  if (SYSTEM_COLOR_VALUES.has(colorValue)) return systemColorDesc;
  if (PRESET_COLOR_VALUES.has(colorValue)) return presetColorDesc;
  if (SCHEME_COLOR_VALUES.has(colorValue)) return schemeColorDesc;
  return rgbColorDesc;
}

// ── SolidFill descriptor ──

export const solidFillDesc: CustomDescriptor<SolidFillOptions> = {
  kind: "custom",
  stringify(color, ctx) {
    const desc = getColorDescriptor(color);
    const inner = stringify(desc, color, ctx);
    if (!inner) return undefined;
    return `<a:solidFill>${inner}</a:solidFill>`;
  },
  parse(el, _ctx) {
    if (!el.elements) return {} as SolidFillOptions;
    for (const child of el.elements) {
      switch (child.name) {
        case "a:srgbClr":
          return rgbColorDesc.parse(child, _ctx);
        case "a:schemeClr":
          return schemeColorDesc.parse(child, _ctx);
        case "a:hslClr":
          return hslColorDesc.parse(child, _ctx);
        case "a:sysClr":
          return systemColorDesc.parse(child, _ctx);
        case "a:prstClr":
          return presetColorDesc.parse(child, _ctx);
        case "a:scrgbClr":
          return scRgbColorDesc.parse(child, _ctx);
      }
    }
    return {} as SolidFillOptions;
  },
};
