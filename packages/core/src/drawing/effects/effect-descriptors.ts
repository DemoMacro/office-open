/**
 * Effect list descriptor for DrawingML shapes.
 *
 * @module
 */

import { escapeXml } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { findChild } from "@office-open/xml";

import type { CustomDescriptor, ReadContext, WriteContext } from "../../descriptor";
import { parse } from "../../descriptor";
import {
  convertToEmu,
  emitAngle,
  emitPercent,
  mapOptional,
  parseAngle,
  parsePercent,
} from "../../util/converters";
import { xsdBlendMode, xsdPresetShadow, xsdRectAlignment } from "../../util/mappings";
import { parseOnOff } from "../../util/values";
import { parseColorChoice, solidFillDesc, stringifyColorChoice } from "../color/color-descriptors";
import type { SolidFillOptions } from "../color/solid-fill";
import { gradientFillDesc, patternFillDesc } from "../fill/fill-descriptors";
import type { BlurEffectOptions, EffectListOptions } from "./effect-list";
import { createFillOverlayEffect } from "./fill-overlay";
import type { FillOverlayEffectOptions } from "./fill-overlay";
import type { GlowEffectOptions } from "./glow";
import type { InnerShadowEffectOptions } from "./inner-shadow";
import type { OuterShadowEffectOptions } from "./outer-shadow";
import type { PresetShadowEffectOptions } from "./preset-shadow";
import type { ReflectionEffectOptions } from "./reflection";

// Parse an ST_Percentage attr that may be a numeric scalar or "50%" literal.
function parsePercentAttr(raw: string | number | undefined): number | undefined {
  if (raw === undefined) return undefined;
  const s = typeof raw === "number" ? String(raw) : raw;
  if (s.endsWith("%")) return Number(s.slice(0, -1));
  return parsePercent(Number(s));
}

// ── Helper: stringify a color into an effect element ──

function stringifyEffectColor(
  color: SolidFillOptions | undefined,
  ctx: WriteContext,
): string | undefined {
  if (!color) return undefined;
  // Effect elements expect EG_ColorChoice (direct color), NOT wrapped in solidFill
  return stringifyColorChoice(color, ctx);
}

function stringifyColorEffect(
  tag: string,
  attrs: Record<string, string | number | undefined>,
  color: SolidFillOptions | undefined,
  ctx: WriteContext,
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

function readColorFromElement(el: XmlElement, ctx: ReadContext): SolidFillOptions | undefined {
  // Effect elements contain EG_ColorChoice directly (not wrapped in solidFill).
  // Reuse parseColorChoice for full coverage (srgb/scheme/hsl/sys/prst/scrgb + transforms).
  const color = parseColorChoice(el, ctx);
  if (!color || Object.keys(color).length === 0) return undefined;
  return color;
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

    // Fill overlay — CT_FillOverlayEffect requires an EG_FillProperties child.
    if (opts.fillOverlay) {
      parts.push(createFillOverlayEffect(opts.fillOverlay));
    }

    // Glow
    if (opts.glow) {
      parts.push(
        stringifyColorEffect(
          "a:glow",
          { rad: mapOptional(opts.glow.radius, convertToEmu) },
          opts.glow.color,
          ctx,
        ) ?? "",
      );
    }

    // Inner shadow
    if (opts.innerShadow) {
      parts.push(
        stringifyColorEffect(
          "a:innerShdw",
          {
            blurRad: mapOptional(opts.innerShadow.blurRadius, convertToEmu),
            dist: mapOptional(opts.innerShadow.distance, convertToEmu),
            dir:
              opts.innerShadow.direction !== undefined
                ? emitAngle(opts.innerShadow.direction)
                : undefined,
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
            blurRad: mapOptional(opts.outerShadow.blurRadius, convertToEmu),
            dist: mapOptional(opts.outerShadow.distance, convertToEmu),
            dir:
              opts.outerShadow.direction !== undefined
                ? emitAngle(opts.outerShadow.direction)
                : undefined,
            sx:
              opts.outerShadow.scaleX !== undefined
                ? emitPercent(opts.outerShadow.scaleX)
                : undefined,
            sy:
              opts.outerShadow.scaleY !== undefined
                ? emitPercent(opts.outerShadow.scaleY)
                : undefined,
            kx:
              opts.outerShadow.skewX !== undefined ? emitAngle(opts.outerShadow.skewX) : undefined,
            ky:
              opts.outerShadow.skewY !== undefined ? emitAngle(opts.outerShadow.skewY) : undefined,
            algn:
              opts.outerShadow.alignment !== undefined
                ? xsdRectAlignment.to(opts.outerShadow.alignment)
                : undefined,
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
            prst: xsdPresetShadow.to(opts.presetShadow.preset),
            dist: mapOptional(opts.presetShadow.distance, convertToEmu),
            dir:
              opts.presetShadow.direction !== undefined
                ? emitAngle(opts.presetShadow.direction)
                : undefined,
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
      if (refOpts.blurRadius !== undefined)
        attrParts.push(`blurRad="${convertToEmu(refOpts.blurRadius)}"`);
      if (refOpts.startAlpha !== undefined)
        attrParts.push(`stA="${emitPercent(refOpts.startAlpha)}"`);
      if (refOpts.startPosition !== undefined)
        attrParts.push(`stPos="${emitPercent(refOpts.startPosition)}"`);
      if (refOpts.endAlpha !== undefined) attrParts.push(`endA="${emitPercent(refOpts.endAlpha)}"`);
      if (refOpts.endPosition !== undefined)
        attrParts.push(`endPos="${emitPercent(refOpts.endPosition)}"`);
      if (refOpts.distance !== undefined)
        attrParts.push(`dist="${convertToEmu(refOpts.distance)}"`);
      if (refOpts.direction !== undefined) attrParts.push(`dir="${emitAngle(refOpts.direction)}"`);
      if (refOpts.fadeDirection !== undefined)
        attrParts.push(`fadeDir="${emitAngle(refOpts.fadeDirection)}"`);
      if (refOpts.scaleX !== undefined) attrParts.push(`sx="${emitPercent(refOpts.scaleX)}"`);
      if (refOpts.scaleY !== undefined) attrParts.push(`sy="${emitPercent(refOpts.scaleY)}"`);
      if (refOpts.skewX !== undefined) attrParts.push(`kx="${emitAngle(refOpts.skewX)}"`);
      if (refOpts.skewY !== undefined) attrParts.push(`ky="${emitAngle(refOpts.skewY)}"`);
      if (refOpts.alignment !== undefined)
        attrParts.push(`algn="${escapeXml(xsdRectAlignment.to(refOpts.alignment))}"`);
      if (refOpts.rotWithShape !== undefined)
        attrParts.push(`rotWithShape="${refOpts.rotWithShape ? 1 : 0}"`);
      const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
      parts.push(`<a:reflection${attrStr}/>`);
    }

    // Soft edge
    if (opts.softEdge !== undefined) {
      parts.push(`<a:softEdge rad="${convertToEmu(opts.softEdge)}"/>`);
    }

    // Filter empty strings
    const content = parts.filter(Boolean).join("");
    if (!content) return undefined;
    return `<a:effectLst>${content}</a:effectLst>`;
  },
  parse(el, ctx) {
    const result: EffectListOptions = {};

    // Blur
    const blur = findChild(el, "a:blur");
    if (blur) {
      const blurOpts: BlurEffectOptions = {};
      if (blur.attributes?.["rad"] !== undefined) blurOpts.radius = Number(blur.attributes["rad"]);
      if (blur.attributes?.["grow"] !== undefined)
        blurOpts.grow = parseOnOff(blur.attributes["grow"]) ?? true;
      result.blur = blurOpts;
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
        innerOpts.direction = parseAngle(Number(innerShdw.attributes["dir"]));
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
        outerOpts.direction = parseAngle(Number(outerShdw.attributes["dir"]));
      if (outerShdw.attributes?.["sx"] !== undefined)
        outerOpts.scaleX = parsePercentAttr(outerShdw.attributes["sx"]);
      if (outerShdw.attributes?.["sy"] !== undefined)
        outerOpts.scaleY = parsePercentAttr(outerShdw.attributes["sy"]);
      if (outerShdw.attributes?.["kx"] !== undefined)
        outerOpts.skewX = parseAngle(Number(outerShdw.attributes["kx"]));
      if (outerShdw.attributes?.["ky"] !== undefined)
        outerOpts.skewY = parseAngle(Number(outerShdw.attributes["ky"]));
      if (outerShdw.attributes?.["algn"] !== undefined)
        outerOpts.alignment = xsdRectAlignment.from(
          String(outerShdw.attributes["algn"]),
        ) as OuterShadowEffectOptions["alignment"];
      if (outerShdw.attributes?.["rotWithShape"] !== undefined)
        outerOpts.rotWithShape = parseOnOff(outerShdw.attributes["rotWithShape"]) ?? true;
      const color = readColorFromElement(outerShdw, ctx);
      if (color) outerOpts.color = color;
      result.outerShadow = outerOpts as OuterShadowEffectOptions;
    }

    // Fill overlay
    const fillOverlay = findChild(el, "a:fillOverlay");
    if (fillOverlay) {
      const overlayOpts: Partial<FillOverlayEffectOptions> = {};
      if (fillOverlay.attributes?.["blend"] !== undefined)
        overlayOpts.blend = xsdBlendMode.from(
          String(fillOverlay.attributes["blend"]),
        ) as FillOverlayEffectOptions["blend"];
      const foSolid = findChild(fillOverlay, "a:solidFill");
      if (foSolid) overlayOpts.solidFill = parse(solidFillDesc, foSolid, ctx) as SolidFillOptions;
      const foGrad = findChild(fillOverlay, "a:gradFill");
      if (foGrad) overlayOpts.gradientFill = parse(gradientFillDesc, foGrad, ctx);
      const foPatt = findChild(fillOverlay, "a:pattFill");
      if (foPatt) overlayOpts.patternFill = parse(patternFillDesc, foPatt, ctx);
      if (findChild(fillOverlay, "a:noFill")) overlayOpts.noFill = true;
      if (findChild(fillOverlay, "a:grpFill")) overlayOpts.groupFill = true;
      result.fillOverlay = overlayOpts as FillOverlayEffectOptions;
    }

    // Preset shadow
    const prstShdw = findChild(el, "a:prstShdw");
    if (prstShdw) {
      const prstOpts: Partial<PresetShadowEffectOptions> = {};
      if (prstShdw.attributes?.["prst"] !== undefined)
        prstOpts.preset = xsdPresetShadow.from(
          String(prstShdw.attributes["prst"]),
        ) as PresetShadowEffectOptions["preset"];
      if (prstShdw.attributes?.["dist"] !== undefined)
        prstOpts.distance = Number(prstShdw.attributes["dist"]);
      if (prstShdw.attributes?.["dir"] !== undefined)
        prstOpts.direction = parseAngle(Number(prstShdw.attributes["dir"]));
      const color = readColorFromElement(prstShdw, ctx);
      if (color) prstOpts.color = color;
      result.presetShadow = prstOpts as PresetShadowEffectOptions;
    }

    // Reflection
    const reflection = findChild(el, "a:reflection");
    if (reflection) {
      const refOpts: ReflectionEffectOptions = {};
      if (reflection.attributes?.["blurRad"] !== undefined)
        refOpts.blurRadius = Number(reflection.attributes["blurRad"]);
      if (reflection.attributes?.["stA"] !== undefined)
        refOpts.startAlpha = parsePercentAttr(reflection.attributes["stA"]);
      if (reflection.attributes?.["stPos"] !== undefined)
        refOpts.startPosition = parsePercentAttr(reflection.attributes["stPos"]);
      if (reflection.attributes?.["endA"] !== undefined)
        refOpts.endAlpha = parsePercentAttr(reflection.attributes["endA"]);
      if (reflection.attributes?.["endPos"] !== undefined)
        refOpts.endPosition = parsePercentAttr(reflection.attributes["endPos"]);
      if (reflection.attributes?.["dist"] !== undefined)
        refOpts.distance = Number(reflection.attributes["dist"]);
      if (reflection.attributes?.["dir"] !== undefined)
        refOpts.direction = parseAngle(Number(reflection.attributes["dir"]));
      if (reflection.attributes?.["fadeDir"] !== undefined)
        refOpts.fadeDirection = parseAngle(Number(reflection.attributes["fadeDir"]));
      if (reflection.attributes?.["sx"] !== undefined)
        refOpts.scaleX = parsePercentAttr(reflection.attributes["sx"]);
      if (reflection.attributes?.["sy"] !== undefined)
        refOpts.scaleY = parsePercentAttr(reflection.attributes["sy"]);
      if (reflection.attributes?.["kx"] !== undefined)
        refOpts.skewX = parseAngle(Number(reflection.attributes["kx"]));
      if (reflection.attributes?.["ky"] !== undefined)
        refOpts.skewY = parseAngle(Number(reflection.attributes["ky"]));
      if (reflection.attributes?.["algn"] !== undefined)
        refOpts.alignment = xsdRectAlignment.from(String(reflection.attributes["algn"]));
      if (reflection.attributes?.["rotWithShape"] !== undefined)
        refOpts.rotWithShape = parseOnOff(reflection.attributes["rotWithShape"]) ?? true;
      result.reflection = refOpts;
    }

    // Soft edge
    const softEdge = findChild(el, "a:softEdge");
    if (softEdge?.attributes?.["rad"] !== undefined) {
      result.softEdge = Number(softEdge.attributes["rad"]);
    }

    return result;
  },
};
