/**
 * Blip descriptor for DrawingML pictures.
 *
 * @module
 */

import { escapeXml } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { findChild } from "@office-open/xml";

import type { CustomDescriptor } from "../../descriptor";
import { stringify, parse } from "../../descriptor";
import { xsdRectAlignment } from "../../util/mappings";
import { solidFillDesc } from "../color/color-descriptors";
import type { SolidFillOptions } from "../color/solid-fill";
import type { BlipOptions } from "./blip";
import type {
  BlipEffectsOptions,
  LuminanceEffectOptions,
  HSLEffectOptions,
  TintEffectOptions,
  BiLevelEffectOptions,
  AlphaReplaceEffectOptions,
  AlphaBiLevelEffectOptions,
  AlphaModulateFixedEffectOptions,
  ColorChangeEffectOptions,
  BlipBlurEffectOptions,
} from "./blip-effects";
import type { BlipFillOptions } from "./blip-fill";
import type { SourceRectangleOptions } from "./source-rectangle";
import type { TileOptions } from "./tile";

// ── Tile descriptor (a:tile) ──

export const tileDesc: CustomDescriptor<TileOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const attrParts: string[] = [];
    if (opts.tx !== undefined) attrParts.push(`tx="${opts.tx}"`);
    if (opts.ty !== undefined) attrParts.push(`ty="${opts.ty}"`);
    if (opts.sx !== undefined) attrParts.push(`sx="${opts.sx}"`);
    if (opts.sy !== undefined) attrParts.push(`sy="${opts.sy}"`);
    if (opts.flip !== undefined) attrParts.push(`flip="${escapeXml(opts.flip)}"`);
    if (opts.align !== undefined)
      attrParts.push(`algn="${escapeXml(xsdRectAlignment.to(opts.align))}"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
    return `<a:tile${attrStr}/>`;
  },
  parse(el, _ctx) {
    const result: Partial<TileOptions> = {};
    if (el.attributes?.["tx"] !== undefined) result.tx = Number(el.attributes["tx"]);
    if (el.attributes?.["ty"] !== undefined) result.ty = Number(el.attributes["ty"]);
    if (el.attributes?.["sx"] !== undefined) result.sx = Number(el.attributes["sx"]);
    if (el.attributes?.["sy"] !== undefined) result.sy = Number(el.attributes["sy"]);
    if (el.attributes?.["flip"] !== undefined)
      result.flip = String(el.attributes["flip"]) as TileOptions["flip"];
    if (el.attributes?.["algn"] !== undefined)
      result.align = xsdRectAlignment.from(String(el.attributes["algn"])) as TileOptions["align"];
    return result as TileOptions;
  },
};

// ── SourceRectangle descriptor (a:srcRect) ──

export const sourceRectangleDesc: CustomDescriptor<SourceRectangleOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const attrParts: string[] = [];
    if (opts.left !== undefined) attrParts.push(`l="${opts.left}"`);
    if (opts.top !== undefined) attrParts.push(`t="${opts.top}"`);
    if (opts.right !== undefined) attrParts.push(`r="${opts.right}"`);
    if (opts.bottom !== undefined) attrParts.push(`b="${opts.bottom}"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
    return `<a:srcRect${attrStr}/>`;
  },
  parse(el, _ctx) {
    const result: Partial<SourceRectangleOptions> = {};
    if (el.attributes?.["l"] !== undefined) result.left = Number(el.attributes["l"]);
    if (el.attributes?.["t"] !== undefined) result.top = Number(el.attributes["t"]);
    if (el.attributes?.["r"] !== undefined) result.right = Number(el.attributes["r"]);
    if (el.attributes?.["b"] !== undefined) result.bottom = Number(el.attributes["b"]);
    return result as SourceRectangleOptions;
  },
};

// ── Stretch descriptor (a:stretch) ──

export const stretchDesc: CustomDescriptor<SourceRectangleOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const attrParts: string[] = [];
    if (opts.left !== undefined) attrParts.push(`l="${opts.left}"`);
    if (opts.top !== undefined) attrParts.push(`t="${opts.top}"`);
    if (opts.right !== undefined) attrParts.push(`r="${opts.right}"`);
    if (opts.bottom !== undefined) attrParts.push(`b="${opts.bottom}"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
    return `<a:stretch><a:fillRect${attrStr}/></a:stretch>`;
  },
  parse(el, _ctx) {
    const fillRect = findChild(el, "a:fillRect");
    if (!fillRect) return {};
    const result: Partial<SourceRectangleOptions> = {};
    if (fillRect.attributes?.["l"] !== undefined) result.left = Number(fillRect.attributes["l"]);
    if (fillRect.attributes?.["t"] !== undefined) result.top = Number(fillRect.attributes["t"]);
    if (fillRect.attributes?.["r"] !== undefined) result.right = Number(fillRect.attributes["r"]);
    if (fillRect.attributes?.["b"] !== undefined) result.bottom = Number(fillRect.attributes["b"]);
    return result as SourceRectangleOptions;
  },
};

// ── Blip effects helpers ──

function stringifyBlipEffects(opts: BlipEffectsOptions, ctx: any): string {
  const parts: string[] = [];

  if (opts.grayscale) {
    parts.push("<a:grayscl/>");
  }

  if (opts.luminance) {
    const attrParts: string[] = [];
    if (opts.luminance.bright !== undefined) attrParts.push(`bright="${opts.luminance.bright}%"`);
    if (opts.luminance.contrast !== undefined)
      attrParts.push(`contrast="${opts.luminance.contrast}%"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
    parts.push(`<a:lum${attrStr}/>`);
  }

  if (opts.hsl) {
    const attrParts: string[] = [];
    if (opts.hsl.hue !== undefined) attrParts.push(`hue="${opts.hsl.hue}"`);
    if (opts.hsl.saturation !== undefined) attrParts.push(`sat="${opts.hsl.saturation}%"`);
    if (opts.hsl.luminance !== undefined) attrParts.push(`lum="${opts.hsl.luminance}%"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
    parts.push(`<a:hsl${attrStr}/>`);
  }

  if (opts.tint) {
    const attrParts: string[] = [];
    if (opts.tint.hue !== undefined) attrParts.push(`hue="${opts.tint.hue}"`);
    if (opts.tint.amount !== undefined) attrParts.push(`amt="${opts.tint.amount}%"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
    parts.push(`<a:tint${attrStr}/>`);
  }

  if (opts.duotone) {
    const c1 = stringify(solidFillDesc, opts.duotone.color1, ctx);
    const c2 = stringify(solidFillDesc, opts.duotone.color2, ctx);
    parts.push(`<a:duotone>${c1 ?? ""}${c2 ?? ""}</a:duotone>`);
  }

  if (opts.biLevel) {
    parts.push(`<a:biLevel thresh="${opts.biLevel.threshold}%"/>`);
  }

  if (opts.alphaCeiling) {
    parts.push("<a:alphaCeiling/>");
  }

  if (opts.alphaFloor) {
    parts.push("<a:alphaFloor/>");
  }

  if (opts.alphaInverse !== undefined) {
    if (typeof opts.alphaInverse === "boolean") {
      parts.push("<a:alphaInv/>");
    } else {
      const colorXml = stringify(solidFillDesc, opts.alphaInverse, ctx);
      parts.push(`<a:alphaInv>${colorXml ?? ""}</a:alphaInv>`);
    }
  }

  if (opts.alphaModFix) {
    const amt = opts.alphaModFix.amount ?? 100;
    parts.push(`<a:alphaModFix amt="${amt}%"/>`);
  }

  if (opts.alphaRepl) {
    parts.push(`<a:alphaRepl a="${opts.alphaRepl.amount}%"/>`);
  }

  if (opts.alphaBiLevel) {
    parts.push(`<a:alphaBiLevel thresh="${opts.alphaBiLevel.threshold}%"/>`);
  }

  if (opts.colorChange) {
    const fromXml = stringify(solidFillDesc, opts.colorChange.from, ctx);
    const toXml = stringify(solidFillDesc, opts.colorChange.to, ctx);
    const attrParts: string[] = [];
    if (opts.colorChange.useAlpha === false) attrParts.push('useA="0"');
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
    parts.push(
      `<a:clrChange${attrStr}><a:clrFrom>${fromXml ?? ""}</a:clrFrom><a:clrTo>${toXml ?? ""}</a:clrTo></a:clrChange>`,
    );
  }

  if (opts.colorRepl) {
    const colorXml = stringify(solidFillDesc, opts.colorRepl.color, ctx);
    parts.push(`<a:clrRepl>${colorXml ?? ""}</a:clrRepl>`);
  }

  if (opts.blur) {
    const attrParts: string[] = [];
    if (opts.blur.radius !== undefined) attrParts.push(`rad="${opts.blur.radius}"`);
    if (opts.blur.grow === false) attrParts.push('grow="0"');
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
    parts.push(`<a:blur${attrStr}/>`);
  }

  return parts.join("");
}

function readBlipEffects(el: XmlElement, ctx: any): BlipEffectsOptions | undefined {
  const result: Partial<BlipEffectsOptions> = {};

  if (findChild(el, "a:grayscl")) result.grayscale = true;

  const lum = findChild(el, "a:lum");
  if (lum) {
    const opts: Partial<LuminanceEffectOptions> = {};
    if (lum.attributes?.["bright"] !== undefined)
      opts.bright = Number(String(lum.attributes["bright"]).replace("%", ""));
    if (lum.attributes?.["contrast"] !== undefined)
      opts.contrast = Number(String(lum.attributes["contrast"]).replace("%", ""));
    result.luminance = opts as LuminanceEffectOptions;
  }

  const hsl = findChild(el, "a:hsl");
  if (hsl) {
    const opts: Partial<HSLEffectOptions> = {};
    if (hsl.attributes?.["hue"] !== undefined) opts.hue = Number(hsl.attributes["hue"]);
    if (hsl.attributes?.["sat"] !== undefined)
      opts.saturation = Number(String(hsl.attributes["sat"]).replace("%", ""));
    if (hsl.attributes?.["lum"] !== undefined)
      opts.luminance = Number(String(hsl.attributes["lum"]).replace("%", ""));
    result.hsl = opts as HSLEffectOptions;
  }

  const tint = findChild(el, "a:tint");
  if (tint) {
    const opts: Partial<TintEffectOptions> = {};
    if (tint.attributes?.["hue"] !== undefined) opts.hue = Number(tint.attributes["hue"]);
    if (tint.attributes?.["amt"] !== undefined)
      opts.amount = Number(String(tint.attributes["amt"]).replace("%", ""));
    result.tint = opts as TintEffectOptions;
  }

  const biLevel = findChild(el, "a:biLevel");
  if (biLevel?.attributes?.["thresh"] !== undefined) {
    result.biLevel = {
      threshold: Number(String(biLevel.attributes["thresh"]).replace("%", "")),
    } as BiLevelEffectOptions;
  }

  if (findChild(el, "a:alphaCeiling")) result.alphaCeiling = true;
  if (findChild(el, "a:alphaFloor")) result.alphaFloor = true;

  const alphaInv = findChild(el, "a:alphaInv");
  if (alphaInv) {
    const solidFill = findChild(alphaInv, "a:solidFill");
    if (solidFill) {
      result.alphaInverse = parse(solidFillDesc, solidFill, ctx) as SolidFillOptions;
    } else {
      result.alphaInverse = {} as SolidFillOptions;
    }
  }

  const alphaModFix = findChild(el, "a:alphaModFix");
  if (alphaModFix) {
    const opts: Partial<AlphaModulateFixedEffectOptions> = {};
    if (alphaModFix.attributes?.["amt"] !== undefined)
      opts.amount = Number(String(alphaModFix.attributes["amt"]).replace("%", ""));
    result.alphaModFix = opts as AlphaModulateFixedEffectOptions;
  }

  const alphaRepl = findChild(el, "a:alphaRepl");
  if (alphaRepl?.attributes?.["a"] !== undefined) {
    result.alphaRepl = {
      amount: Number(String(alphaRepl.attributes["a"]).replace("%", "")),
    } as AlphaReplaceEffectOptions;
  }

  const alphaBiLevel = findChild(el, "a:alphaBiLevel");
  if (alphaBiLevel?.attributes?.["thresh"] !== undefined) {
    result.alphaBiLevel = {
      threshold: Number(String(alphaBiLevel.attributes["thresh"]).replace("%", "")),
    } as AlphaBiLevelEffectOptions;
  }

  const clrChange = findChild(el, "a:clrChange");
  if (clrChange) {
    const opts: Partial<ColorChangeEffectOptions> = {};
    if (clrChange.attributes?.["useA"] !== undefined)
      opts.useAlpha = clrChange.attributes["useA"] !== "0";
    const clrFrom = findChild(clrChange, "a:clrFrom");
    if (clrFrom) {
      const fromFill = findChild(clrFrom, "a:solidFill");
      if (fromFill) opts.from = parse(solidFillDesc, fromFill, ctx) as SolidFillOptions;
    }
    const clrTo = findChild(clrChange, "a:clrTo");
    if (clrTo) {
      const toFill = findChild(clrTo, "a:solidFill");
      if (toFill) opts.to = parse(solidFillDesc, toFill, ctx) as SolidFillOptions;
    }
    result.colorChange = opts as ColorChangeEffectOptions;
  }

  const clrRepl = findChild(el, "a:clrRepl");
  if (clrRepl) {
    const solidFill = findChild(clrRepl, "a:solidFill");
    if (solidFill)
      result.colorRepl = { color: parse(solidFillDesc, solidFill, ctx) as SolidFillOptions };
  }

  const blur = findChild(el, "a:blur");
  if (blur) {
    const opts: Partial<BlipBlurEffectOptions> = {};
    if (blur.attributes?.["rad"] !== undefined) opts.radius = Number(blur.attributes["rad"]);
    if (blur.attributes?.["grow"] !== undefined) opts.grow = blur.attributes["grow"] !== "0";
    result.blur = opts as BlipBlurEffectOptions;
  }

  const duotone = findChild(el, "a:duotone");
  if (duotone?.elements) {
    // Try to read two solidFill children
    const fills: SolidFillOptions[] = [];
    for (const child of duotone.elements) {
      const sf = findChild(child, "a:solidFill");
      if (sf) fills.push(parse(solidFillDesc, sf, ctx) as SolidFillOptions);
    }
    if (fills.length >= 2) {
      result.duotone = { color1: fills[0], color2: fills[1] };
    }
  }

  return Object.keys(result).length > 0 ? (result as BlipEffectsOptions) : undefined;
}

// ── Blip descriptor (a:blip) ──

export const blipDesc: CustomDescriptor<BlipOptions & { blipEffects?: BlipEffectsOptions }> = {
  kind: "custom",
  stringify(opts, ctx) {
    const attrParts: string[] = [];
    const embedValue = `{${opts.referenceId}}`;
    attrParts.push(`r:embed="${escapeXml(embedValue)}"`);
    attrParts.push('cstate="none"');
    const attrStr = " " + attrParts.join(" ");

    const parts: string[] = [];
    if (opts.blipEffects) {
      parts.push(stringifyBlipEffects(opts.blipEffects, ctx));
    }

    const content = parts.join("");
    if (!content) return `<a:blip${attrStr}/>`;
    return `<a:blip${attrStr}>${content}</a:blip>`;
  },
  parse(el, ctx) {
    const result: Partial<BlipOptions & { blipEffects?: BlipEffectsOptions }> = {};
    const embed = el.attributes?.["r:embed"];
    if (embed !== undefined) {
      // Strip { and } wrapper if present
      result.referenceId = String(embed).replace(/^\{(.+)\}$/, "$1");
    }
    const link = el.attributes?.["r:link"];
    if (link !== undefined) {
      result.referenceId = String(link).replace(/^\{(.+)\}$/, "$1");
    }
    const effects = readBlipEffects(el, ctx);
    if (effects) result.blipEffects = effects;
    return result as BlipOptions & { blipEffects?: BlipEffectsOptions };
  },
};

// ── BlipFill descriptor (pic:blipFill / a:blipFill) ──

export const blipFillDesc: CustomDescriptor<
  BlipFillOptions & { referenceId?: string; blipEffects?: BlipEffectsOptions }
> = {
  kind: "custom",
  stringify(opts, ctx) {
    const attrParts: string[] = [];
    if (opts.dpi !== undefined) attrParts.push(`dpi="${opts.dpi}"`);
    if (opts.rotWithShape !== undefined)
      attrParts.push(`rotWithShape="${opts.rotWithShape ? 1 : 0}"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";

    const parts: string[] = [];

    // Blip child (uses referenceId from parent)
    if (opts.referenceId) {
      const blipOpts = { referenceId: opts.referenceId, blipEffects: opts.blipEffects };
      const blipXml = stringify(blipDesc, blipOpts, ctx);
      if (blipXml) parts.push(blipXml);
    }

    // Source rectangle
    if (opts.srcRect) {
      const srcRectXml = stringify(sourceRectangleDesc, opts.srcRect, ctx);
      if (srcRectXml) parts.push(srcRectXml);
    }

    // Tile or stretch
    if (opts.tile) {
      const tileXml = stringify(tileDesc, opts.tile, ctx);
      if (tileXml) parts.push(tileXml);
    } else {
      parts.push("<a:stretch><a:fillRect/></a:stretch>");
    }

    const content = parts.join("");
    if (!attrStr && !content) return undefined;
    if (!content) return `<pic:blipFill${attrStr}/>`;
    return `<pic:blipFill${attrStr}>${content}</pic:blipFill>`;
  },
  parse(el, ctx) {
    const result: Partial<
      BlipFillOptions & { referenceId?: string; blipEffects?: BlipEffectsOptions }
    > = {};

    // Attributes
    if (el.attributes?.["dpi"] !== undefined) result.dpi = Number(el.attributes["dpi"]);
    if (el.attributes?.["rotWithShape"] !== undefined)
      result.rotWithShape = el.attributes["rotWithShape"] !== "0";

    // Blip child
    const blip = findChild(el, "a:blip");
    if (blip) {
      const blipResult = parse(blipDesc, blip, ctx);
      if (blipResult.referenceId) result.referenceId = blipResult.referenceId;
      if (blipResult.blipEffects) result.blipEffects = blipResult.blipEffects;
    }

    // Source rectangle
    const srcRect = findChild(el, "a:srcRect");
    if (srcRect)
      result.srcRect = parse(sourceRectangleDesc, srcRect, ctx) as SourceRectangleOptions;

    // Tile
    const tile = findChild(el, "a:tile");
    if (tile) result.tile = parse(tileDesc, tile, ctx) as TileOptions;

    return result as BlipFillOptions & { referenceId?: string; blipEffects?: BlipEffectsOptions };
  },
};
