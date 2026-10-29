import { findChild, attr } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { ReadContext } from "../../descriptor";
import { parse } from "../../descriptor";
import { parseAngle, parsePercentAttr } from "../../util/converters";
import { xsdBlendMode, xsdEffectContainer } from "../../util/mappings";
import { parseOnOff } from "../../util/values";
import type {
  AlphaBiLevelEffectOptions,
  AlphaModulateFixedEffectOptions,
  AlphaReplaceEffectOptions,
  ColorChangeEffectOptions,
  DuotoneEffectOptions,
  HSLEffectOptions,
  LuminanceEffectOptions,
  TintEffectOptions,
} from "../blip/blip-effects";
import { parseColorChoice } from "../color/color-descriptors";
import type { SolidFillOptions } from "../color/solid-fill";
import { fillDesc, findFillChild } from "../fill/fill-descriptors";
import type { EffectDagOptions } from "./effect-dag";
import type {
  AlphaInverseEffectOptions,
  AlphaOutsetEffectOptions,
  BlendEffectOptions,
  EffectReferenceOptions,
  FillEffectOptions,
  RelativeOffsetEffectOptions,
  TransformEffectOptions,
} from "./effect-dag";
import { effectListDesc } from "./effect-descriptors";

function parseColor(element: Element, ctx: ReadContext): SolidFillOptions | undefined {
  return parseColorChoice(element, ctx);
}

function parseRequiredColor(element: Element, ctx: ReadContext): SolidFillOptions {
  const color = parseColor(element, ctx);
  if (!color) throw new Error(`Missing color in ${element.name}.`);
  return color;
}

function parseRequiredPercent(value: string | undefined): number {
  return parsePercentAttr(value ?? "0") ?? 0;
}

function parseColorChildren(element: Element, ctx: ReadContext): SolidFillOptions[] {
  return (element.elements ?? [])
    .filter((child) => child.type === "element")
    .map((child) => parseColor(child, ctx))
    .filter((color): color is SolidFillOptions => color !== undefined);
}

function parseFillEffect(element: Element, ctx: ReadContext): FillEffectOptions {
  const fillElement = findFillChild(element);
  if (!fillElement) throw new Error("a:fill requires exactly one fill child.");
  return { fill: parse(fillDesc, fillElement, ctx) };
}

function parseDuotone(element: Element, ctx: ReadContext): DuotoneEffectOptions {
  const [color1, color2] = parseColorChildren(element, ctx);
  if (!color1 || !color2) throw new Error("dgm effect duotone requires two colors.");
  return {
    color1,
    color2,
    ...(attr(element, "xfr1") !== undefined ? { transformReference1: attr(element, "xfr1") } : {}),
    ...(attr(element, "xfr2") !== undefined ? { transformReference2: attr(element, "xfr2") } : {}),
  };
}

function parseEffectChild(element: Element, ctx: ReadContext, result: EffectDagOptions): void {
  switch (element.name) {
    case "a:blur":
    case "a:fillOverlay":
    case "a:glow":
    case "a:innerShdw":
    case "a:outerShdw":
    case "a:prstShdw":
    case "a:reflection":
    case "a:softEdge": {
      const parsed = effectListDesc.parse(
        { type: "element", name: "a:effectLst", elements: [element] },
        ctx,
      );
      Object.assign(result, parsed);
      return;
    }
    case "a:cont":
      result.containers = [...(result.containers ?? []), parseEffectContainer(element, ctx)];
      return;
    case "a:effect": {
      const reference: EffectReferenceOptions = { ref: attr(element, "ref") ?? "" };
      result.effectRefs = [...(result.effectRefs ?? []), reference];
      return;
    }
    case "a:alphaBiLevel": {
      const effect: AlphaBiLevelEffectOptions = {
        threshold: parseRequiredPercent(attr(element, "thresh")),
      };
      result.alphaBiLevel = effect;
      return;
    }
    case "a:alphaCeiling":
      result.alphaCeiling = true;
      return;
    case "a:alphaFloor":
      result.alphaFloor = true;
      return;
    case "a:alphaInv": {
      const effect: AlphaInverseEffectOptions = {};
      const color = (element.elements ?? []).find((child) => child.type === "element");
      if (color) effect.color = parseRequiredColor(color, ctx);
      result.alphaInverse = effect;
      return;
    }
    case "a:alphaMod": {
      const nested = (element.elements ?? []).find((child) => child.type === "element");
      if (nested) result.alphaModulate = parseEffectContainer(nested, ctx);
      return;
    }
    case "a:alphaModFix": {
      const effect: AlphaModulateFixedEffectOptions = {};
      const amount = attr(element, "amt");
      if (amount !== undefined) effect.amount = parsePercentAttr(amount);
      result.alphaModulateFixed = effect;
      return;
    }
    case "a:alphaOutset": {
      const effect: AlphaOutsetEffectOptions = {};
      const radius = attr(element, "rad");
      if (radius !== undefined) effect.radius = Number(radius);
      result.alphaOutset = effect;
      return;
    }
    case "a:alphaRepl": {
      const effect: AlphaReplaceEffectOptions = {
        alpha: parseRequiredPercent(attr(element, "a")),
      };
      result.alphaReplace = effect;
      return;
    }
    case "a:biLevel":
      result.biLevel = { threshold: parseRequiredPercent(attr(element, "thresh")) };
      return;
    case "a:blend": {
      const nested = findChild(element, "a:cont");
      if (!nested) throw new Error("a:blend requires a nested effect container.");
      const effect: BlendEffectOptions = {
        blend: xsdBlendMode.from(attr(element, "blend") ?? "") as BlendEffectOptions["blend"],
        container: parseEffectContainer(nested, ctx),
      };
      result.blend = effect;
      return;
    }
    case "a:clrChange": {
      const from = findChild(element, "a:clrFrom");
      const to = findChild(element, "a:clrTo");
      if (!from || !to) throw new Error("a:clrChange requires clrFrom and clrTo.");
      const effect: ColorChangeEffectOptions = {
        from: parseRequiredColor(from, ctx),
        to: parseRequiredColor(to, ctx),
      };
      const useAlpha = parseOnOff(attr(element, "useA"));
      if (useAlpha === false) effect.useAlpha = false;
      result.colorChange = effect;
      return;
    }
    case "a:clrRepl": {
      const color = (element.elements ?? []).find((child) => child.type === "element");
      if (color) result.colorReplace = parseRequiredColor(color, ctx);
      return;
    }
    case "a:duotone":
      result.duotone = parseDuotone(element, ctx);
      return;
    case "a:fill":
      result.fill = parseFillEffect(element, ctx);
      return;
    case "a:grayscl":
      result.grayscale = true;
      return;
    case "a:hsl": {
      const effect: HSLEffectOptions = {};
      const hue = attr(element, "hue");
      const saturation = attr(element, "sat");
      const luminance = attr(element, "lum");
      if (hue !== undefined) effect.hue = parseAngle(Number(hue));
      if (saturation !== undefined) effect.saturation = parsePercentAttr(saturation);
      if (luminance !== undefined) effect.luminance = parsePercentAttr(luminance);
      result.hsl = effect;
      return;
    }
    case "a:lum": {
      const effect: LuminanceEffectOptions = {};
      const bright = attr(element, "bright");
      const contrast = attr(element, "contrast");
      if (bright !== undefined) effect.bright = parsePercentAttr(bright);
      if (contrast !== undefined) effect.contrast = parsePercentAttr(contrast);
      result.luminance = effect;
      return;
    }
    case "a:tint": {
      const effect: TintEffectOptions = {};
      const hue = attr(element, "hue");
      const amount = attr(element, "amt");
      if (hue !== undefined) effect.hue = parseAngle(Number(hue));
      if (amount !== undefined) effect.amount = parsePercentAttr(amount);
      result.tint = effect;
      return;
    }
    case "a:relOff": {
      const effect: RelativeOffsetEffectOptions = {};
      const translateX = attr(element, "tx");
      const translateY = attr(element, "ty");
      if (translateX !== undefined) effect.translateX = parsePercentAttr(translateX);
      if (translateY !== undefined) effect.translateY = parsePercentAttr(translateY);
      result.relativeOffset = effect;
      return;
    }
    case "a:xfrm": {
      const effect: TransformEffectOptions = {};
      const scaleX = attr(element, "sx");
      const scaleY = attr(element, "sy");
      const skewX = attr(element, "kx");
      const skewY = attr(element, "ky");
      const translateX = attr(element, "tx");
      const translateY = attr(element, "ty");
      if (scaleX !== undefined) effect.scaleX = parsePercentAttr(scaleX);
      if (scaleY !== undefined) effect.scaleY = parsePercentAttr(scaleY);
      if (skewX !== undefined) effect.skewX = parseAngle(Number(skewX));
      if (skewY !== undefined) effect.skewY = parseAngle(Number(skewY));
      if (translateX !== undefined) effect.translateX = parsePercentAttr(translateX);
      if (translateY !== undefined) effect.translateY = parsePercentAttr(translateY);
      result.transform = effect;
      return;
    }
    default:
      throw new Error(`Unsupported effect container child: ${element.name}.`);
  }
}

export function parseEffectContainer(element: Element, ctx: ReadContext): EffectDagOptions {
  const result: EffectDagOptions = {};
  const type = attr(element, "type");
  const name = attr(element, "name");
  if (type) result.type = xsdEffectContainer.from(type) as EffectDagOptions["type"];
  if (name !== undefined) result.name = name;
  for (const child of element.elements ?? []) {
    if (child.type === "element") parseEffectChild(child, ctx, result);
  }
  return result;
}
