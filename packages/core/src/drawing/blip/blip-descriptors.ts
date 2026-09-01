/**
 * Blip descriptor for DrawingML pictures.
 *
 * @module
 */

import { escapeXml } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { findChild, attr, stringifyElement } from "@office-open/xml";

import type { CustomDescriptor, ReadContext, WriteContext } from "../../descriptor";
import { stringify, parse } from "../../descriptor";
import { emitAngle, emitPercent, parseAngle, parsePercentAttr } from "../../util/converters";
import { extUriMatches } from "../../util/ext-uri";
import { xsdRectAlignment } from "../../util/mappings";
import { parseOnOff } from "../../util/values";
import {
  parseColorChoice,
  parseColorChoiceElement,
  stringifyColorChoice,
} from "../color/color-descriptors";
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
    if (opts.sx !== undefined) attrParts.push(`sx="${emitPercent(opts.sx)}"`);
    if (opts.sy !== undefined) attrParts.push(`sy="${emitPercent(opts.sy)}"`);
    if (opts.flip !== undefined) attrParts.push(`flip="${escapeXml(opts.flip)}"`);
    if (opts.alignment !== undefined)
      attrParts.push(`algn="${escapeXml(xsdRectAlignment.to(opts.alignment))}"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
    return `<a:tile${attrStr}/>`;
  },
  parse(el, _ctx) {
    const result: TileOptions = {};
    if (el.attributes?.["tx"] !== undefined) result.tx = Number(el.attributes["tx"]);
    if (el.attributes?.["ty"] !== undefined) result.ty = Number(el.attributes["ty"]);
    if (el.attributes?.["sx"] !== undefined) result.sx = parsePercentAttr(el.attributes["sx"])!;
    if (el.attributes?.["sy"] !== undefined) result.sy = parsePercentAttr(el.attributes["sy"])!;
    if (el.attributes?.["flip"] !== undefined)
      result.flip = String(el.attributes["flip"]) as TileOptions["flip"];
    if (el.attributes?.["algn"] !== undefined)
      result.alignment = xsdRectAlignment.from(
        String(el.attributes["algn"]),
      ) as TileOptions["alignment"];
    return result;
  },
};

// ── SourceRectangle descriptor (a:srcRect) ──

export const sourceRectangleDesc: CustomDescriptor<SourceRectangleOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const attrParts: string[] = [];
    if (opts.left !== undefined) attrParts.push(`l="${emitPercent(opts.left)}"`);
    if (opts.top !== undefined) attrParts.push(`t="${emitPercent(opts.top)}"`);
    if (opts.right !== undefined) attrParts.push(`r="${emitPercent(opts.right)}"`);
    if (opts.bottom !== undefined) attrParts.push(`b="${emitPercent(opts.bottom)}"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
    return `<a:srcRect${attrStr}/>`;
  },
  parse(el, _ctx) {
    const result: SourceRectangleOptions = {};
    if (el.attributes?.["l"] !== undefined) result.left = parsePercentAttr(el.attributes["l"])!;
    if (el.attributes?.["t"] !== undefined) result.top = parsePercentAttr(el.attributes["t"])!;
    if (el.attributes?.["r"] !== undefined) result.right = parsePercentAttr(el.attributes["r"])!;
    if (el.attributes?.["b"] !== undefined) result.bottom = parsePercentAttr(el.attributes["b"])!;
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

/** Serialize a:blip adjustment effects (a:lum, a:duotone, …) in public units. */
export function stringifyBlipEffects(opts: BlipEffectsOptions, ctx: WriteContext): string {
  const parts: string[] = [];

  if (opts.grayscale) {
    parts.push("<a:grayscl/>");
  }

  if (opts.luminance) {
    const attrParts: string[] = [];
    if (opts.luminance.bright !== undefined)
      attrParts.push(`bright="${emitPercent(opts.luminance.bright)}"`);
    if (opts.luminance.contrast !== undefined)
      attrParts.push(`contrast="${emitPercent(opts.luminance.contrast)}"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
    parts.push(`<a:lum${attrStr}/>`);
  }

  if (opts.hsl) {
    const attrParts: string[] = [];
    if (opts.hsl.hue !== undefined) attrParts.push(`hue="${emitAngle(opts.hsl.hue)}"`);
    if (opts.hsl.saturation !== undefined)
      attrParts.push(`sat="${emitPercent(opts.hsl.saturation)}"`);
    if (opts.hsl.luminance !== undefined)
      attrParts.push(`lum="${emitPercent(opts.hsl.luminance)}"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
    parts.push(`<a:hsl${attrStr}/>`);
  }

  if (opts.tint) {
    const attrParts: string[] = [];
    if (opts.tint.hue !== undefined) attrParts.push(`hue="${emitAngle(opts.tint.hue)}"`);
    if (opts.tint.amount !== undefined) attrParts.push(`amt="${emitPercent(opts.tint.amount)}"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
    parts.push(`<a:tint${attrStr}/>`);
  }

  if (opts.duotone) {
    const c1 = stringifyColorChoice(opts.duotone.color1, ctx);
    const c2 = stringifyColorChoice(opts.duotone.color2, ctx);
    parts.push(`<a:duotone>${c1}${c2}</a:duotone>`);
  }

  if (opts.biLevel) {
    parts.push(`<a:biLevel thresh="${emitPercent(opts.biLevel.threshold)}"/>`);
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
      parts.push(`<a:alphaInv>${stringifyColorChoice(opts.alphaInverse, ctx)}</a:alphaInv>`);
    }
  }

  if (opts.alphaModulateFixed) {
    const amt = emitPercent(opts.alphaModulateFixed.amount ?? 100);
    parts.push(`<a:alphaModFix amt="${amt}"/>`);
  }

  if (opts.alphaReplace) {
    parts.push(`<a:alphaRepl a="${emitPercent(opts.alphaReplace.alpha)}"/>`);
  }

  if (opts.alphaBiLevel) {
    parts.push(`<a:alphaBiLevel thresh="${emitPercent(opts.alphaBiLevel.threshold)}"/>`);
  }

  if (opts.colorChange) {
    const fromXml = stringifyColorChoice(opts.colorChange.from, ctx);
    const toXml = stringifyColorChoice(opts.colorChange.to, ctx);
    const attrParts: string[] = [];
    if (opts.colorChange.useAlpha === false) attrParts.push('useA="0"');
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";
    parts.push(
      `<a:clrChange${attrStr}><a:clrFrom>${fromXml ?? ""}</a:clrFrom><a:clrTo>${toXml ?? ""}</a:clrTo></a:clrChange>`,
    );
  }

  if (opts.colorReplace) {
    parts.push(`<a:clrRepl>${stringifyColorChoice(opts.colorReplace.color, ctx)}</a:clrRepl>`);
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

function readBlipEffects(el: XmlElement, ctx: ReadContext): BlipEffectsOptions | undefined {
  const result: BlipEffectsOptions = {};

  if (findChild(el, "a:grayscl")) result.grayscale = true;

  const lum = findChild(el, "a:lum");
  if (lum) {
    const opts: LuminanceEffectOptions = {};
    const bright = parsePercentAttr(lum.attributes?.["bright"]);
    if (bright !== undefined) opts.bright = bright;
    const contrast = parsePercentAttr(lum.attributes?.["contrast"]);
    if (contrast !== undefined) opts.contrast = contrast;
    result.luminance = opts;
  }

  const hsl = findChild(el, "a:hsl");
  if (hsl) {
    const opts: HSLEffectOptions = {};
    if (hsl.attributes?.["hue"] !== undefined) opts.hue = parseAngle(Number(hsl.attributes["hue"]));
    const sat = parsePercentAttr(hsl.attributes?.["sat"]);
    if (sat !== undefined) opts.saturation = sat;
    const l = parsePercentAttr(hsl.attributes?.["lum"]);
    if (l !== undefined) opts.luminance = l;
    result.hsl = opts;
  }

  const tint = findChild(el, "a:tint");
  if (tint) {
    const opts: TintEffectOptions = {};
    if (tint.attributes?.["hue"] !== undefined)
      opts.hue = parseAngle(Number(tint.attributes["hue"]));
    const amt = parsePercentAttr(tint.attributes?.["amt"]);
    if (amt !== undefined) opts.amount = amt;
    result.tint = opts;
  }

  const biLevel = findChild(el, "a:biLevel");
  if (biLevel?.attributes?.["thresh"] !== undefined) {
    result.biLevel = { threshold: parsePercentAttr(biLevel.attributes["thresh"])! };
  }

  if (findChild(el, "a:alphaCeiling")) result.alphaCeiling = true;
  if (findChild(el, "a:alphaFloor")) result.alphaFloor = true;

  const alphaInv = findChild(el, "a:alphaInv");
  if (alphaInv) {
    const color = parseColorChoice(alphaInv, ctx);
    result.alphaInverse = Object.keys(color).length > 0 ? color : ({} as SolidFillOptions);
  }

  const alphaModFix = findChild(el, "a:alphaModFix");
  if (alphaModFix) {
    const opts: AlphaModulateFixedEffectOptions = {};
    const amt = parsePercentAttr(alphaModFix.attributes?.["amt"]);
    if (amt !== undefined) opts.amount = amt;
    result.alphaModulateFixed = opts;
  }

  const alphaRepl = findChild(el, "a:alphaRepl");
  if (alphaRepl?.attributes?.["a"] !== undefined) {
    result.alphaReplace = { alpha: parsePercentAttr(alphaRepl.attributes["a"])! };
  }

  const alphaBiLevel = findChild(el, "a:alphaBiLevel");
  if (alphaBiLevel?.attributes?.["thresh"] !== undefined) {
    result.alphaBiLevel = { threshold: parsePercentAttr(alphaBiLevel.attributes["thresh"])! };
  }

  const clrChange = findChild(el, "a:clrChange");
  if (clrChange) {
    const opts: Partial<ColorChangeEffectOptions> = {};
    if (clrChange.attributes?.["useA"] !== undefined)
      opts.useAlpha = parseOnOff(clrChange.attributes["useA"]) ?? true;
    const clrFrom = findChild(clrChange, "a:clrFrom");
    if (clrFrom) opts.from = parseColorChoice(clrFrom, ctx);
    const clrTo = findChild(clrChange, "a:clrTo");
    if (clrTo) opts.to = parseColorChoice(clrTo, ctx);
    result.colorChange = opts as ColorChangeEffectOptions;
  }

  const clrRepl = findChild(el, "a:clrRepl");
  if (clrRepl) {
    result.colorReplace = { color: parseColorChoice(clrRepl, ctx) as SolidFillOptions };
  }

  const blur = findChild(el, "a:blur");
  if (blur) {
    const opts: BlipBlurEffectOptions = {};
    if (blur.attributes?.["rad"] !== undefined) opts.radius = Number(blur.attributes["rad"]);
    if (blur.attributes?.["grow"] !== undefined)
      opts.grow = parseOnOff(blur.attributes["grow"]) ?? true;
    result.blur = opts;
  }

  const duotone = findChild(el, "a:duotone");
  if (duotone?.elements) {
    // CT_Duotone carries two bare EG_ColorChoice children (no solidFill
    // wrapper) — take the first two color elements in order.
    const fills: SolidFillOptions[] = [];
    for (const child of duotone.elements) {
      const color = parseColorChoiceElement(child, ctx);
      if (color) fills.push(color);
    }
    const [color1, color2] = fills;
    if (color1 && color2) {
      result.duotone = { color1, color2 };
    }
  }

  return Object.keys(result).length > 0 ? (result as BlipEffectsOptions) : undefined;
}

// ── Blip descriptor (a:blip) ──

/** The a14:useLocalDpi extension uri (CT_Blip extLst). */
const USE_LOCAL_DPI_EXT_URI = "{28A0092B-C50C-407E-A947-70E740481C1C}";

/** Input shape of blipDesc — BlipOptions plus round-trip-only blip content. */
export type BlipDescriptorOptions = BlipOptions & {
  blipEffects?: BlipEffectsOptions;
  /**
   * a14:useLocalDpi from the blip extension list — Office's local-DPI display
   * hint, the dominant non-SVG a:blip ext content.
   */
  useLocalDpi?: boolean;
  /**
   * Verbatim `a:extLst` inner XML for extensions beyond useLocalDpi (a14
   * imgProps artistic effects, …). Round-trip only: takes precedence over
   * {@link useLocalDpi}, which it subsumes when the source list carries both.
   */
  ext?: string;
};

export const blipDesc: CustomDescriptor<BlipDescriptorOptions> = {
  kind: "custom",
  stringify(opts, ctx) {
    // Attribute order matches Office output: r:embed, cstate, r:link. Every
    // attribute is optional (CT_Blip) and omitted when unset — a linked-only
    // picture carries r:link alone.
    const attrParts: string[] = [];
    if (opts.referenceId !== undefined)
      attrParts.push(`r:embed="{${escapeXml(opts.referenceId)}}"`);
    if (opts.compression !== undefined) attrParts.push(`cstate="${opts.compression}"`);
    if (opts.linkReferenceId !== undefined)
      attrParts.push(`r:link="{${escapeXml(opts.linkReferenceId)}}"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";

    const parts: string[] = [];
    if (opts.blipEffects) {
      parts.push(stringifyBlipEffects(opts.blipEffects, ctx));
    }
    // The verbatim channel subsumes useLocalDpi when both would apply.
    if (opts.ext !== undefined) {
      parts.push(`<a:extLst>${opts.ext}</a:extLst>`);
    } else if (opts.useLocalDpi !== undefined) {
      parts.push(
        `<a:extLst><a:ext uri="${USE_LOCAL_DPI_EXT_URI}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="${opts.useLocalDpi ? 1 : 0}"/></a:ext></a:extLst>`,
      );
    }

    const content = parts.join("");
    if (!content) return `<a:blip${attrStr}/>`;
    return `<a:blip${attrStr}>${content}</a:blip>`;
  },
  parse(el, ctx) {
    const result: Partial<BlipDescriptorOptions> = {};
    const embed = el.attributes?.["r:embed"];
    if (embed !== undefined) {
      // Strip { and } wrapper if present
      result.referenceId = String(embed).replace(/^\{(.+)\}$/, "$1");
    }
    const link = el.attributes?.["r:link"];
    if (link !== undefined) {
      result.linkReferenceId = String(link).replace(/^\{(.+)\}$/, "$1");
    }
    const cstate = el.attributes?.["cstate"];
    if (cstate !== undefined) result.compression = cstate as BlipDescriptorOptions["compression"];
    const effects = readBlipEffects(el, ctx);
    if (effects) result.blipEffects = effects;
    const extLst = findChild(el, "a:extLst");
    if (extLst) {
      // Extensions beyond useLocalDpi (a14 imgProps effects, …) have no
      // structured model — keep the whole list verbatim and skip the
      // useLocalDpi extraction so the two channels never double-emit.
      const hasUnmodeled = (extLst.elements ?? []).some(
        (ext) => ext.name === "a:ext" && !extUriMatches(attr(ext, "uri"), USE_LOCAL_DPI_EXT_URI),
      );
      if (hasUnmodeled) {
        const inner = (extLst.elements ?? []).map((e) => stringifyElement(e)).join("");
        if (inner) result.ext = inner;
      } else {
        for (const ext of extLst.elements ?? []) {
          if (ext.name !== "a:ext" || !extUriMatches(attr(ext, "uri"), USE_LOCAL_DPI_EXT_URI))
            continue;
          const useLocalDpi = findChild(ext, "a14:useLocalDpi");
          if (useLocalDpi !== undefined)
            result.useLocalDpi = parseOnOff(attr(useLocalDpi, "val")) ?? true;
        }
      }
    }
    return result as BlipDescriptorOptions;
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
      const blipOpts = {
        referenceId: opts.referenceId,
        compression: opts.compression,
        blipEffects: opts.blipEffects,
      };
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
      result.rotWithShape = parseOnOff(el.attributes["rotWithShape"]) ?? true;

    // Blip child
    const blip = findChild(el, "a:blip");
    if (blip) {
      const blipResult = parse(blipDesc, blip, ctx);
      if (blipResult.referenceId) result.referenceId = blipResult.referenceId;
      if (blipResult.compression !== undefined) result.compression = blipResult.compression;
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
