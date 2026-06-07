/**
 * Effect list descriptor for DrawingML shapes.
 *
 * @module
 */

import { escapeXml } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { findChild } from "@office-open/xml";

import type { CustomDescriptor } from "../../descriptor";
import { stringify, parse } from "../../descriptor";
import { rgbColorDesc, schemeColorDesc } from "../color/color-descriptors";
import type { SolidFillOptions } from "../color/solid-fill";
import type { BlurEffectOptions, EffectListOptions } from "./effect-list";
import type { GlowEffectOptions } from "./glow";
import type { InnerShadowEffectOptions } from "./inner-shadow";
import type { OuterShadowEffectOptions } from "./outer-shadow";
import type { ReflectionEffectOptions } from "./reflection";

// ── Helper: stringify a color into an effect element ──

function stringifyEffectColor(color: SolidFillOptions | undefined, ctx: any): string | undefined {
  if (!color) return undefined;
  // Effect elements expect EG_ColorChoice (direct color), NOT wrapped in solidFill
  const desc = getColorDesc(color);
  return stringify(desc, color, ctx);
}

function getColorDesc(color: SolidFillOptions): CustomDescriptor<any> {
  if ("hue" in color && "saturation" in color && "luminance" in color) {
    // HSL — not imported here but rare in effects; fall through to rgbColorDesc
  }
  if ("r" in color && "g" in color && "b" in color) {
    // scRGB — fall through to rgbColorDesc for simplicity
  }
  // Check if it looks like a scheme color
  const val = (color as { value: string }).value;
  if (val && !/^[0-9a-fA-F]{6}$/.test(val)) {
    // Not a hex RGB — likely scheme/system/preset color
    return schemeColorDesc;
  }
  return rgbColorDesc;
}

function stringifyColorEffect(
  tag: string,
  attrs: Record<string, string | number | undefined>,
  color: SolidFillOptions | undefined,
  ctx: any,
): string | undefined {
  const attrParts: string[] = [];
  for (const [key, val] of Object.entries(attrs)) {
    if (val !== undefined) attrParts.push(`${key}="${escapeXml(String(val as number | string))}"`);
  }
  const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";

  const colorXml = stringifyEffectColor(color, ctx);
  if (!colorXml && !attrStr) return undefined;

  if (!colorXml) return `<${tag}${attrStr}/>`;
  return `<${tag}${attrStr}>${colorXml}</${tag}>`;
}

function readColorFromElement(el: XmlElement, ctx: any): SolidFillOptions | undefined {
  // Effect elements contain EG_ColorChoice directly (not wrapped in solidFill)
  for (const child of el.elements ?? []) {
    switch (child.name) {
      case "a:srgbClr":
        return rgbColorDesc.parse(child, ctx) as SolidFillOptions;
      case "a:schemeClr":
        return parse(schemeColorDesc, child, ctx) as SolidFillOptions;
    }
  }
  return undefined;
}

// ── EffectList descriptor ──

export const effectListDesc: CustomDescriptor<EffectListOptions> = {
  kind: "custom",
  stringify(opts, ctx) {
    const parts: string[] = [];

    // Blur
    if (opts.blur) {
      const attrParts: string[] = [];
      if (opts.blur.radius !== undefined) attrParts.push(`rad="${opts.blur.radius}"`);
      if (opts.blur.grow === false) attrParts.push('grow="0"');
      const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
      parts.push(`<a:blur${attrStr}/>`);
    }

    // Fill overlay
    if (opts.fillOverlay) {
      parts.push(`<a:fillOverlay blend="${escapeXml(opts.fillOverlay.blend)}"/>`);
    }

    // Glow
    if (opts.glow) {
      parts.push(
        stringifyColorEffect("a:glow", { rad: opts.glow.radius }, opts.glow.color, ctx) ?? "",
      );
    }

    // Inner shadow
    if (opts.innerShadow) {
      parts.push(
        stringifyColorEffect(
          "a:innerShdw",
          {
            blurRad: opts.innerShadow.blurRadius,
            dist: opts.innerShadow.distance,
            dir: opts.innerShadow.direction,
          },
          opts.innerShadow.color,
          ctx,
        ) ?? "",
      );
    }

    // Outer shadow
    if (opts.outerShadow) {
      parts.push(
        stringifyColorEffect(
          "a:outerShdw",
          {
            blurRad: opts.outerShadow.blurRadius,
            dist: opts.outerShadow.distance,
            dir: opts.outerShadow.direction,
            sx: opts.outerShadow.scaleX,
            sy: opts.outerShadow.scaleY,
            kx: opts.outerShadow.skewX,
            ky: opts.outerShadow.skewY,
            algn: opts.outerShadow.alignment,
            rotWithShape: opts.outerShadow.rotWithShape === false ? 0 : undefined,
          },
          opts.outerShadow.color,
          ctx,
        ) ?? "",
      );
    }

    // Preset shadow
    if (opts.presetShadow) {
      parts.push(
        stringifyColorEffect(
          "a:prstShdw",
          {
            prst: opts.presetShadow.preset,
            dist: opts.presetShadow.distance,
            dir: opts.presetShadow.direction,
          },
          opts.presetShadow.color,
          ctx,
        ) ?? "",
      );
    }

    // Reflection
    if (opts.reflection) {
      const refOpts = opts.reflection === true ? ({} as ReflectionEffectOptions) : opts.reflection;
      const attrParts: string[] = [];
      if (refOpts.blurRadius !== undefined) attrParts.push(`blurRad="${refOpts.blurRadius}"`);
      if (refOpts.startAlpha !== undefined) attrParts.push(`stA="${refOpts.startAlpha}"`);
      if (refOpts.startPosition !== undefined) attrParts.push(`stPos="${refOpts.startPosition}"`);
      if (refOpts.endAlpha !== undefined) attrParts.push(`endA="${refOpts.endAlpha}"`);
      if (refOpts.endPosition !== undefined) attrParts.push(`endPos="${refOpts.endPosition}"`);
      if (refOpts.distance !== undefined) attrParts.push(`dist="${refOpts.distance}"`);
      if (refOpts.direction !== undefined) attrParts.push(`dir="${refOpts.direction}"`);
      if (refOpts.fadeDirection !== undefined) attrParts.push(`fadeDir="${refOpts.fadeDirection}"`);
      if (refOpts.scaleX !== undefined) attrParts.push(`sx="${refOpts.scaleX}"`);
      if (refOpts.scaleY !== undefined) attrParts.push(`sy="${refOpts.scaleY}"`);
      if (refOpts.skewX !== undefined) attrParts.push(`kx="${refOpts.skewX}"`);
      if (refOpts.skewY !== undefined) attrParts.push(`ky="${refOpts.skewY}"`);
      if (refOpts.alignment !== undefined) attrParts.push(`algn="${refOpts.alignment}"`);
      if (refOpts.rotWithShape !== undefined)
        attrParts.push(`rotWithShape="${refOpts.rotWithShape ? 1 : 0}"`);
      const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
      parts.push(`<a:reflection${attrStr}/>`);
    }

    // Soft edge
    if (opts.softEdge !== undefined) {
      parts.push(`<a:softEdge rad="${opts.softEdge}"/>`);
    }

    // Filter empty strings
    const content = parts.filter(Boolean).join("");
    if (!content) return undefined;
    return `<a:effectLst>${content}</a:effectLst>`;
  },
  parse(el, ctx) {
    const result: Partial<EffectListOptions> = {};

    // Blur
    const blur = findChild(el, "a:blur");
    if (blur) {
      const blurOpts: Partial<BlurEffectOptions> = {};
      if (blur.attributes?.["rad"] !== undefined) blurOpts.radius = Number(blur.attributes["rad"]);
      if (blur.attributes?.["grow"] !== undefined) blurOpts.grow = blur.attributes["grow"] !== "0";
      result.blur = blurOpts as BlurEffectOptions;
    }

    // Glow
    const glow = findChild(el, "a:glow");
    if (glow) {
      const glowOpts: Partial<GlowEffectOptions> = {};
      if (glow.attributes?.["rad"] !== undefined) glowOpts.radius = Number(glow.attributes["rad"]);
      const color = readColorFromElement(glow, ctx);
      if (color) glowOpts.color = color;
      result.glow = glowOpts as GlowEffectOptions;
    }

    // Inner shadow
    const innerShdw = findChild(el, "a:innerShdw");
    if (innerShdw) {
      const innerOpts: Partial<InnerShadowEffectOptions> = {};
      if (innerShdw.attributes?.["blurRad"] !== undefined)
        innerOpts.blurRadius = Number(innerShdw.attributes["blurRad"]);
      if (innerShdw.attributes?.["dist"] !== undefined)
        innerOpts.distance = Number(innerShdw.attributes["dist"]);
      if (innerShdw.attributes?.["dir"] !== undefined)
        innerOpts.direction = Number(innerShdw.attributes["dir"]);
      const color = readColorFromElement(innerShdw, ctx);
      if (color) innerOpts.color = color;
      result.innerShadow = innerOpts as InnerShadowEffectOptions;
    }

    // Outer shadow
    const outerShdw = findChild(el, "a:outerShdw");
    if (outerShdw) {
      const outerOpts: Partial<OuterShadowEffectOptions> = {};
      if (outerShdw.attributes?.["blurRad"] !== undefined)
        outerOpts.blurRadius = Number(outerShdw.attributes["blurRad"]);
      if (outerShdw.attributes?.["dist"] !== undefined)
        outerOpts.distance = Number(outerShdw.attributes["dist"]);
      if (outerShdw.attributes?.["dir"] !== undefined)
        outerOpts.direction = Number(outerShdw.attributes["dir"]);
      if (outerShdw.attributes?.["sx"] !== undefined)
        outerOpts.scaleX = Number(outerShdw.attributes["sx"]);
      if (outerShdw.attributes?.["sy"] !== undefined)
        outerOpts.scaleY = Number(outerShdw.attributes["sy"]);
      if (outerShdw.attributes?.["kx"] !== undefined)
        outerOpts.skewX = Number(outerShdw.attributes["kx"]);
      if (outerShdw.attributes?.["ky"] !== undefined)
        outerOpts.skewY = Number(outerShdw.attributes["ky"]);
      if (outerShdw.attributes?.["algn"] !== undefined)
        outerOpts.alignment = String(
          outerShdw.attributes["algn"],
        ) as OuterShadowEffectOptions["alignment"];
      if (outerShdw.attributes?.["rotWithShape"] !== undefined)
        outerOpts.rotWithShape = outerShdw.attributes["rotWithShape"] !== "0";
      const color = readColorFromElement(outerShdw, ctx);
      if (color) outerOpts.color = color;
      result.outerShadow = outerOpts as OuterShadowEffectOptions;
    }

    // Soft edge
    const softEdge = findChild(el, "a:softEdge");
    if (softEdge?.attributes?.["rad"] !== undefined) {
      result.softEdge = Number(softEdge.attributes["rad"]);
    }

    return result;
  },
};
