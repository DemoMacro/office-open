/**
 * Blip descriptor for DrawingML pictures.
 *
 * @module
 */

import { escapeXml } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { findChild } from "@office-open/xml";

import type { CustomDescriptor, ReadContext, WriteContext } from "../../descriptor";
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
    if (opts.sx !== undefined) attrParts.push(`sx="${Math.round(opts.sx * 1000)}"`);
    if (opts.sy !== undefined) attrParts.push(`sy="${Math.round(opts.sy * 1000)}"`);
    if (opts.flip !== undefined) attrParts.push(`flip="${escapeXml(opts.flip)}"`);
    if (opts.align !== undefined)
      attrParts.push(`algn="${escapeXml(xsdRectAlignment.to(opts.align))}"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
    return `<a:tile${attrStr}/>`;
  },
  parse(el, _ctx) {
    const result: TileOptions = {};
    if (el.attributes?.["tx"] !== undefined) result.tx = Number(el.attributes["tx"]);
    if (el.attributes?.["ty"] !== undefined) result.ty = Number(el.attributes["ty"]);
    if (el.attributes?.["sx"] !== undefined) result.sx = parsePercent(el.attributes["sx"])!;
    if (el.attributes?.["sy"] !== undefined) result.sy = parsePercent(el.attributes["sy"])!;
    if (el.attributes?.["flip"] !== undefined)
      result.flip = String(el.attributes["flip"]) as TileOptions["flip"];
    if (el.attributes?.["algn"] !== undefined)
      result.align = xsdRectAlignment.from(String(el.attributes["algn"])) as TileOptions["align"];
    return result;
  },
};

// ── SourceRectangle descriptor (a:srcRect) ──

export const sourceRectangleDesc: CustomDescriptor<SourceRectangleOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const attrParts: string[] = [];
    if (opts.left !== undefined) attrParts.push(`l="${Math.round(opts.left * 1000)}"`);
    if (opts.top !== undefined) attrParts.push(`t="${Math.round(opts.top * 1000)}"`);
    if (opts.right !== undefined) attrParts.push(`r="${Math.round(opts.right * 1000)}"`);
    if (opts.bottom !== undefined) attrParts.push(`b="${Math.round(opts.bottom * 1000)}"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
    return `<a:srcRect${attrStr}/>`;
  },
  parse(el, _ctx) {
    const result: SourceRectangleOptions = {};
    if (el.attributes?.["l"] !== undefined) result.left = parsePercent(el.attributes["l"])!;
    if (el.attributes?.["t"] !== undefined) result.top = parsePercent(el.attributes["t"])!;
    if (el.attributes?.["r"] !== undefined) result.right = parsePercent(el.attributes["r"])!;
    if (el.attributes?.["b"] !== undefined) result.bottom = parsePercent(el.attributes["b"])!;
    return result;
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
    const result: SourceRectangleOptions = {};
    if (fillRect.attributes?.["l"] !== undefined) result.left = Number(fillRect.attributes["l"]);
    if (fillRect.attributes?.["t"] !== undefined) result.top = Number(fillRect.attributes["t"]);
    if (fillRect.attributes?.["r"] !== undefined) result.right = Number(fillRect.attributes["r"]);
    if (fillRect.attributes?.["b"] !== undefined) result.bottom = Number(fillRect.attributes["b"]);
    return result;
  },
};

// ── Blip effects helpers ──

function stringifyBlipEffects(opts: BlipEffectsOptions, ctx: WriteContext): string {
  const parts: string[] = [];

  if (opts.grayscale) {
    parts.push("<a:grayscl/>");
  }

  if (opts.luminance) {
    const attrParts: string[] = [];
    if (opts.luminance.bright !== undefined)
      attrParts.push(`bright="${Math.round(opts.luminance.bright * 1000)}"`);
    if (opts.luminance.contrast !== undefined)
      attrParts.push(`contrast="${Math.round(opts.luminance.contrast * 1000)}"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
    parts.push(`<a:lum${attrStr}/>`);
  }

  if (opts.hsl) {
    const attrParts: string[] = [];
    if (opts.hsl.hue !== undefined) attrParts.push(`hue="${Math.round(opts.hsl.hue * 60000)}"`);
    if (opts.hsl.saturation !== undefined)
      attrParts.push(`sat="${Math.round(opts.hsl.saturation * 1000)}"`);
    if (opts.hsl.luminance !== undefined)
      attrParts.push(`lum="${Math.round(opts.hsl.luminance * 1000)}"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
    parts.push(`<a:hsl${attrStr}/>`);
  }

  if (opts.tint) {
    const attrParts: string[] = [];
    if (opts.tint.hue !== undefined) attrParts.push(`hue="${Math.round(opts.tint.hue * 60000)}"`);
    if (opts.tint.amount !== undefined)
      attrParts.push(`amt="${Math.round(opts.tint.amount * 1000)}"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
    parts.push(`<a:tint${attrStr}/>`);
  }

  if (opts.duotone) {
    const c1 = stringify(solidFillDesc, opts.duotone.color1, ctx);
    const c2 = stringify(solidFillDesc, opts.duotone.color2, ctx);
    parts.push(`<a:duotone>${c1 ?? ""}${c2 ?? ""}</a:duotone>`);
  }

  if (opts.biLevel) {
    parts.push(`<a:biLevel thresh="${Math.round(opts.biLevel.threshold * 1000)}"/>`);
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
    const amt = Math.round((opts.alphaModFix.amount ?? 100) * 1000);
    parts.push(`<a:alphaModFix amt="${amt}"/>`);
  }

  if (opts.alphaRepl) {
    parts.push(`<a:alphaRepl a="${Math.round(opts.alphaRepl.amount * 1000)}"/>`);
  }

  if (opts.alphaBiLevel) {
    parts.push(`<a:alphaBiLevel thresh="${Math.round(opts.alphaBiLevel.threshold * 1000)}"/>`);
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

// Parse an ST_Percentage attribute that may be a 1/1000th integer ("50000")
// or a percent literal ("50%"); both are XSD-valid. Returns integer percent.
function parsePercent(raw: string | number | undefined): number | undefined {
  if (raw === undefined) return undefined;
  const s = typeof raw === "number" ? String(raw) : raw;
  if (s.endsWith("%")) return Number(s.slice(0, -1));
  return Number(s) / 1000;
}

function readBlipEffects(el: XmlElement, ctx: ReadContext): BlipEffectsOptions | undefined {
  const result: BlipEffectsOptions = {};

  if (findChild(el, "a:grayscl")) result.grayscale = true;

  const lum = findChild(el, "a:lum");
  if (lum) {
    const opts: LuminanceEffectOptions = {};
    const bright = parsePercent(lum.attributes?.["bright"]);
    if (bright !== undefined) opts.bright = bright;
    const contrast = parsePercent(lum.attributes?.["contrast"]);
    if (contrast !== undefined) opts.contrast = contrast;
    result.luminance = opts;
  }

  const hsl = findChild(el, "a:hsl");
  if (hsl) {
    const opts: HSLEffectOptions = {};
    if (hsl.attributes?.["hue"] !== undefined) opts.hue = Number(hsl.attributes["hue"]) / 60000;
    const sat = parsePercent(hsl.attributes?.["sat"]);
    if (sat !== undefined) opts.saturation = sat;
    const l = parsePercent(hsl.attributes?.["lum"]);
    if (l !== undefined) opts.luminance = l;
    result.hsl = opts;
  }

  const tint = findChild(el, "a:tint");
  if (tint) {
    const opts: TintEffectOptions = {};
    if (tint.attributes?.["hue"] !== undefined) opts.hue = Number(tint.attributes["hue"]) / 60000;
    const amt = parsePercent(tint.attributes?.["amt"]);
    if (amt !== undefined) opts.amount = amt;
    result.tint = opts;
  }

  const biLevel = findChild(el, "a:biLevel");
  if (biLevel?.attributes?.["thresh"] !== undefined) {
    result.biLevel = { threshold: parsePercent(biLevel.attributes["thresh"])! };
  }

  if (findChild(el, "a:alphaCeiling")) result.alphaCeiling = true;
  if (findChild(el, "a:alphaFloor")) result.alphaFloor = true;

  const alphaInv = findChild(el, "a:alphaInv");
  if (alphaInv) {
    const solidFill = findChild(alphaInv, "a:solidFill");
    if (solidFill) {
      result.alphaInverse = parse(solidFillDesc, solidFill, ctx);
    } else {
      result.alphaInverse = {} as SolidFillOptions;
    }
  }

  const alphaModFix = findChild(el, "a:alphaModFix");
  if (alphaModFix) {
    const opts: AlphaModulateFixedEffectOptions = {};
    const amt = parsePercent(alphaModFix.attributes?.["amt"]);
    if (amt !== undefined) opts.amount = amt;
    result.alphaModFix = opts;
  }

  const alphaRepl = findChild(el, "a:alphaRepl");
  if (alphaRepl?.attributes?.["a"] !== undefined) {
    result.alphaRepl = { amount: parsePercent(alphaRepl.attributes["a"])! };
  }

  const alphaBiLevel = findChild(el, "a:alphaBiLevel");
  if (alphaBiLevel?.attributes?.["thresh"] !== undefined) {
    result.alphaBiLevel = { threshold: parsePercent(alphaBiLevel.attributes["thresh"])! };
  }

  const clrChange = findChild(el, "a:clrChange");
  if (clrChange) {
    const opts: Partial<ColorChangeEffectOptions> = {};
    if (clrChange.attributes?.["useA"] !== undefined)
      opts.useAlpha = clrChange.attributes["useA"] !== "0";
    const clrFrom = findChild(clrChange, "a:clrFrom");
    if (clrFrom) {
      const fromFill = findChild(clrFrom, "a:solidFill");
      if (fromFill) opts.from = parse(solidFillDesc, fromFill, ctx);
    }
    const clrTo = findChild(clrChange, "a:clrTo");
    if (clrTo) {
      const toFill = findChild(clrTo, "a:solidFill");
      if (toFill) opts.to = parse(solidFillDesc, toFill, ctx);
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
    const opts: BlipBlurEffectOptions = {};
    if (blur.attributes?.["rad"] !== undefined) opts.radius = Number(blur.attributes["rad"]);
    if (blur.attributes?.["grow"] !== undefined) opts.grow = blur.attributes["grow"] !== "0";
    result.blur = opts;
  }

  const duotone = findChild(el, "a:duotone");
  if (duotone?.elements) {
    // Try to read two solidFill children
    const fills: SolidFillOptions[] = [];
    for (const child of duotone.elements) {
      const sf = findChild(child, "a:solidFill");
      if (sf) fills.push(parse(solidFillDesc, sf, ctx) as SolidFillOptions);
    }
    const [color1, color2] = fills;
    if (color1 && color2) {
      result.duotone = { color1, color2 };
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
    if (opts.sourceRectangle) {
      const srcRectXml = stringify(sourceRectangleDesc, opts.sourceRectangle, ctx);
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
    if (srcRect) result.sourceRectangle = parse(sourceRectangleDesc, srcRect, ctx);

    // Tile
    const tile = findChild(el, "a:tile");
    if (tile) result.tile = parse(tileDesc, tile, ctx);

    return result as BlipFillOptions & { referenceId?: string; blipEffects?: BlipEffectsOptions };
  },
};
