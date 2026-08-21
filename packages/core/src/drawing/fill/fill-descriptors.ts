/**
 * Fill descriptors for DrawingML EG_FillProperties.
 *
 * @module
 */

import { element, escapeXml } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { findChild } from "@office-open/xml";

import type { CustomDescriptor, ReadContext } from "../../descriptor";
import { parse } from "../../descriptor";
import { emitAngle, emitPercent, parseAngle, parsePercent } from "../../util/converters";
import { toUint8Array } from "../../util/data-type";
import { uniqueId } from "../../util/generators";
import { imageTypeFromPath } from "../../util/image-type";
import { xsdPattern } from "../../util/mappings";
import { parseOnOff, stripColorHashPrefix } from "../../util/values";
import { blipFillDesc } from "../blip/blip-descriptors";
import { createBlipEffects } from "../blip/blip-effects";
import { createSourceRectangle } from "../blip/source-rectangle";
import { createTileInfo } from "../blip/tile";
import { solidFillDesc, parseColorChoice, emitColorChoice } from "../color/color-descriptors";
import type { SolidFillOptions } from "../color/solid-fill";
import type { BlipFillConfigOptions, FillOptions } from "./fill-options";
import type {
  GradientFillOptions,
  GradientShadeOptions,
  LinearShadeOptions,
  PathShadeOptions,
  RelativeRect,
} from "./gradient-fill";
import type { PatternFillOptions } from "./pattern-fill";

// ── Gradient fill helpers ──

function stringifyRelativeRect(tag: string, rect: RelativeRect): string {
  const parts: string[] = [];
  if (rect.left) parts.push(`l="${escapeXml(rect.left)}"`);
  if (rect.top) parts.push(`t="${escapeXml(rect.top)}"`);
  if (rect.right) parts.push(`r="${escapeXml(rect.right)}"`);
  if (rect.bottom) parts.push(`b="${escapeXml(rect.bottom)}"`);
  const attrStr = parts.length ? " " + parts.join(" ") : "";
  return `<${tag}${attrStr}/>`;
}

function readRelativeRect(el: XmlElement): RelativeRect {
  const result: RelativeRect = {};
  if (el.attributes?.["l"]) result.left = String(el.attributes["l"]);
  if (el.attributes?.["t"]) result.top = String(el.attributes["t"]);
  if (el.attributes?.["r"]) result.right = String(el.attributes["r"]);
  if (el.attributes?.["b"]) result.bottom = String(el.attributes["b"]);
  return result as RelativeRect;
}

function stringifyShade(shade: GradientShadeOptions): string {
  if ("angle" in shade) {
    const parts: string[] = [];
    if (shade.angle !== undefined) parts.push(`ang="${emitAngle(shade.angle)}"`);
    if (shade.scaled !== undefined) parts.push(`scaled="${shade.scaled ? 1 : 0}"`);
    const attrStr = parts.length ? " " + parts.join(" ") : "";
    return `<a:lin${attrStr}/>`;
  }
  const pathShade = shade as PathShadeOptions;
  const parts: string[] = [];
  if (pathShade.path) parts.push(`path="${escapeXml(pathShade.path)}"`);
  const attrStr = parts.length ? " " + parts.join(" ") : "";
  if (pathShade.fillToRectangle) {
    return `<a:path${attrStr}>${stringifyRelativeRect("a:fillToRect", pathShade.fillToRectangle)}</a:path>`;
  }
  return `<a:path${attrStr}/>`;
}

// ── GradientFill descriptor ──

export const gradientFillDesc: CustomDescriptor<GradientFillOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    return emitGradientFillXml(opts);
  },
  parse(el, ctx) {
    const result: Partial<GradientFillOptions> = {};

    // Stops — a:gs contains EG_ColorChoice directly (no solidFill wrapper)
    const gsLst = findChild(el, "a:gsLst");
    if (gsLst?.elements) {
      result.stops = gsLst.elements
        .filter((c) => c.name === "a:gs")
        .map((gs) => {
          const pos = Number(gs.attributes?.["pos"] ?? 0);
          const color = readDirectColor(gs, ctx);
          return { position: parsePercent(pos), color };
        });
    }

    // Shade (a:lin or a:path)
    const lin = findChild(el, "a:lin");
    if (lin) {
      const shade: LinearShadeOptions = {};
      if (lin.attributes?.["ang"] !== undefined)
        shade.angle = parseAngle(Number(lin.attributes["ang"]));
      if (lin.attributes?.["scaled"] !== undefined)
        shade.scaled = parseOnOff(lin.attributes["scaled"]) ?? true;
      result.shade = shade;
    } else {
      const path = findChild(el, "a:path");
      if (path) {
        const shade: PathShadeOptions = {};
        if (path.attributes?.["path"] !== undefined)
          shade.path = String(path.attributes["path"]) as PathShadeOptions["path"];
        const fillToRectangle = findChild(path, "a:fillToRect");
        if (fillToRectangle) shade.fillToRectangle = readRelativeRect(fillToRectangle);
        result.shade = shade;
      }
    }

    // Flip
    if (el.attributes?.["flip"] !== undefined)
      result.flip = String(el.attributes["flip"]) as GradientFillOptions["flip"];

    // Rotate with shape
    if (el.attributes?.["rotWithShape"] !== undefined)
      result.rotateWithShape = parseOnOff(el.attributes["rotWithShape"]) ?? true;

    // Tile rect
    const tileRectangle = findChild(el, "a:tileRect");
    if (tileRectangle) result.tileRectangle = readRelativeRect(tileRectangle);

    return result as GradientFillOptions;
  },
};

// ── PatternFill descriptor ──

export const patternFillDesc: CustomDescriptor<PatternFillOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    return emitPatternFillXml(opts);
  },
  parse(el, ctx) {
    const result: Partial<PatternFillOptions> = {};

    const prst = el.attributes?.["prst"];
    if (prst) result.pattern = xsdPattern.from(String(prst)) as PatternFillOptions["pattern"];

    const fgClr = findChild(el, "a:fgClr");
    if (fgClr) {
      result.foregroundColor = readDirectColor(fgClr, ctx);
    }

    const bgClr = findChild(el, "a:bgClr");
    if (bgClr) {
      result.backgroundColor = readDirectColor(bgClr, ctx);
    }

    return result as PatternFillOptions;
  },
};

// ── Fill (EG_FillProperties) descriptor ──

// ── Fill (EG_FillProperties) serialization — single source ──

/**
 * Serialize any FillOptions (string shorthand or object union, blip included)
 * without a write context. The blip variant takes the embed placeholder
 * directly — callers with a write context get it from `ctx.addMedia`.
 */
export function emitFillXml(options: FillOptions, embedPlaceholder?: string): string {
  // String shorthand → solid fill
  if (typeof options === "string") {
    return `<a:solidFill>${emitColorChoice({ value: stripColorHashPrefix(options) } as SolidFillOptions)}</a:solidFill>`;
  }

  switch (options.type) {
    case "none":
      return "<a:noFill/>";

    case "solid": {
      const color =
        typeof options.color === "string"
          ? ({ value: stripColorHashPrefix(options.color) } as SolidFillOptions)
          : options.color;
      return `<a:solidFill>${emitColorChoice(color)}</a:solidFill>`;
    }

    case "gradient": {
      // Core API variant
      if ("options" in options) {
        return emitGradientFillXml(options.options);
      }
      // Simplified API variant
      const gradOpts: GradientFillOptions = {
        stops: options.stops.map((stop) => ({
          position: stop.position,
          color:
            typeof stop.color === "string"
              ? ({ value: stripColorHashPrefix(stop.color) } as SolidFillOptions)
              : stop.color,
        })),
      };
      if (!options.path && options.angle !== undefined) {
        gradOpts.shade = { angle: options.angle, scaled: options.scaled ?? true };
      }
      if (options.path) {
        gradOpts.shade = { path: options.path };
      }
      return emitGradientFillXml(gradOpts);
    }

    case "blip":
      return emitBlipFill(options, embedPlaceholder);

    case "pattern": {
      const patternOpts: PatternFillOptions = {
        pattern: options.pattern as PatternFillOptions["pattern"],
        ...(options.foregroundColor && {
          foregroundColor:
            typeof options.foregroundColor === "string"
              ? ({ value: stripColorHashPrefix(options.foregroundColor) } as SolidFillOptions)
              : options.foregroundColor,
        }),
        ...(options.backgroundColor && {
          backgroundColor:
            typeof options.backgroundColor === "string"
              ? ({ value: stripColorHashPrefix(options.backgroundColor) } as SolidFillOptions)
              : options.backgroundColor,
        }),
      };
      return emitPatternFillXml(patternOpts);
    }

    case "group":
      return "<a:grpFill/>";
  }
}

/** Serialize a:gradFill from full GradientFillOptions (descriptor emission). */
function emitGradientFillXml(opts: GradientFillOptions): string {
  const parts: string[] = [];

  // Gradient stop list — a:gs expects EG_ColorChoice (direct color), NOT solidFill
  const stopsXml = opts.stops
    .map((stop) => {
      const colorXml = emitColorChoice(stop.color);
      if (!colorXml) return `<a:gs pos="${emitPercent(stop.position)}"/>`;
      return `<a:gs pos="${emitPercent(stop.position)}">${colorXml}</a:gs>`;
    })
    .join("");
  parts.push(`<a:gsLst>${stopsXml}</a:gsLst>`);

  // Shade
  if (opts.shade) parts.push(stringifyShade(opts.shade));

  // Tile rect
  if (opts.tileRectangle) parts.push(stringifyRelativeRect("a:tileRect", opts.tileRectangle));

  // Attributes
  const attrParts: string[] = [];
  if (opts.flip) attrParts.push(`flip="${escapeXml(opts.flip)}"`);
  if (opts.rotateWithShape !== undefined)
    attrParts.push(`rotWithShape="${opts.rotateWithShape ? 1 : 0}"`);
  const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";

  return `<a:gradFill${attrStr}>${parts.join("")}</a:gradFill>`;
}

/** Serialize a:pattFill from PatternFillOptions (descriptor emission). */
function emitPatternFillXml(opts: PatternFillOptions): string {
  const parts: string[] = [];
  const prst = xsdPattern.to(opts.pattern);

  // a:fgClr/a:bgClr expect EG_ColorChoice (direct color), NOT solidFill
  if (opts.foregroundColor) {
    const colorXml = emitColorChoice(opts.foregroundColor);
    if (colorXml) parts.push(`<a:fgClr>${colorXml}</a:fgClr>`);
  }
  if (opts.backgroundColor) {
    const colorXml = emitColorChoice(opts.backgroundColor);
    if (colorXml) parts.push(`<a:bgClr>${colorXml}</a:bgClr>`);
  }

  const inner = parts.join("");
  return `<a:pattFill prst="${escapeXml(prst)}">${inner}</a:pattFill>`;
}

/** Serialize a:blipFill with the given embed reference (descriptor emission). */
function emitBlipFill(options: BlipFillConfigOptions & { type: "blip" }, embed?: string): string {
  // Build a:blip with {fileName} placeholder — the packer's ImageReplacer
  // replaces `{fileName}` with `rId{N}` and creates the relationship. When the
  // caller supplies embed (a media reference already registered with the write
  // context, e.g. `{image1.png}`), use it verbatim so the emitted reference
  // matches the registration.
  const fileName = `${uniqueId()}.${options.imageType ?? "png"}`;
  const embedRef = embed ?? `{${fileName}}`;

  const blipChildren: string[] = [];
  if (options.blipEffects) {
    blipChildren.push(...createBlipEffects(options.blipEffects));
  }
  // noEmbed: the source blip carried no r:embed (an empty marker) — re-emit
  // it bare instead of fabricating a media reference.
  // The r: prefix is declared inline on the blip: hosts whose root declares
  // only the a: namespace (a theme's fmtScheme) still get well-formed XML.
  const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const blip = options.noEmbed
    ? element("a:blip", undefined, blipChildren.length > 0 ? blipChildren : undefined)
    : element(
        "a:blip",
        { cstate: "none", "xmlns:r": R_NS, "r:embed": embedRef },
        blipChildren.length > 0 ? blipChildren : undefined,
      );

  const children: string[] = [blip];
  // a:srcRect is optional in CT_BlipFillProperties — emit it only when the
  // source carried one (round-trip); a bare <a:srcRect/> round-trips as {}.
  if (options.sourceRectangle !== undefined)
    children.push(createSourceRectangle(options.sourceRectangle));
  if (options.tile) {
    children.push(createTileInfo(options.tile));
  } else {
    children.push("<a:stretch><a:fillRect/></a:stretch>");
  }
  const attrs: Record<string, string | number | undefined> = {};
  if (options.dpi !== undefined) attrs.dpi = options.dpi;
  if (options.rotWithShape !== undefined) attrs.rotWithShape = options.rotWithShape ? 1 : 0;
  return element("a:blipFill", attrs, children);
}

export const fillDesc: CustomDescriptor<FillOptions> = {
  kind: "custom",
  stringify(opts, ctx) {
    if (typeof opts !== "string" && opts.type === "blip") {
      // noEmbed: an empty-marker blip — nothing to register with the media
      // store; emit the bare a:blipFill shape (attrs, srcRect, stretch).
      if (opts.noEmbed) return emitBlipFill(opts);
      // Register the image media via the write context, then emit a:blipFill
      // with the returned {fileName} placeholder. The format-package compiler
      // replaces the placeholder with a relationship rId at pack time.
      const placeholder = ctx.addMedia(
        toUint8Array(opts.data!, { encoding: "base64" }),
        opts.imageType!,
      );
      return emitBlipFill(opts, placeholder);
    }
    return emitFillXml(opts);
  },
  parse(el, ctx) {
    // Resolve fill element — either el itself or its first fill child
    const fillEl = el.name !== undefined && FILL_TAGS.has(el.name) ? el : findFillChild(el);
    if (!fillEl) return { type: "none" };

    switch (fillEl.name) {
      case "a:noFill":
        return { type: "none" };

      case "a:solidFill":
        return { type: "solid", color: parse(solidFillDesc, fillEl, ctx) };

      case "a:gradFill":
        return {
          type: "gradient",
          options: parse(gradientFillDesc, fillEl, ctx) as GradientFillOptions,
        };

      case "a:pattFill":
        return { type: "pattern", ...(parse(patternFillDesc, fillEl, ctx) as PatternFillOptions) };

      case "a:grpFill":
        return { type: "group" };

      case "a:blipFill": {
        // Blip fill (image) — resolve r:embed to binary media via the read context
        const blipOpts = parse(blipFillDesc, fillEl, ctx);
        const mediaPath = blipOpts.referenceId
          ? ctx.resolveRelationship(blipOpts.referenceId)
          : undefined;
        const data = mediaPath ? ctx.getRaw(mediaPath) : undefined;
        if (mediaPath && data) {
          const blip: BlipFillConfigOptions & { type: "blip" } = {
            type: "blip",
            data,
            imageType: imageTypeFromPath(mediaPath),
            fileName: mediaPath.split("/").pop(),
          };
          if (blipOpts.dpi !== undefined) blip.dpi = blipOpts.dpi;
          if (blipOpts.rotWithShape !== undefined) blip.rotWithShape = blipOpts.rotWithShape;
          if (blipOpts.blipEffects) blip.blipEffects = blipOpts.blipEffects;
          if (blipOpts.sourceRectangle) blip.sourceRectangle = blipOpts.sourceRectangle;
          if (blipOpts.tile) blip.tile = blipOpts.tile;
          return blip;
        }
        if (blipOpts.referenceId === undefined) {
          // Empty a:blip (no r:embed) — Word's pic:spPr duplicate of
          // pic:blipFill references no image of its own. Keep the fill shape
          // (attrs, srcRect, stretch) instead of degrading to noFill.
          const blip: BlipFillConfigOptions & { type: "blip" } = { type: "blip", noEmbed: true };
          if (blipOpts.dpi !== undefined) blip.dpi = blipOpts.dpi;
          if (blipOpts.rotWithShape !== undefined) blip.rotWithShape = blipOpts.rotWithShape;
          if (blipOpts.sourceRectangle) blip.sourceRectangle = blipOpts.sourceRectangle;
          if (blipOpts.tile) blip.tile = blipOpts.tile;
          return blip;
        }
        return { type: "none" };
      }

      default:
        return { type: "none" };
    }
  },
};

// ── Fill child discovery ──

const FILL_TAGS: ReadonlySet<string> = new Set([
  "a:noFill",
  "a:solidFill",
  "a:gradFill",
  "a:pattFill",
  "a:grpFill",
  "a:blipFill",
]);

/**
 * Find the first EG_FillProperties child of an element (any of the six fill
 * kinds) in a single scan — the shared existence probe for sites that only
 * parse a fill when one is present.
 */
export function findFillChild(el: XmlElement): XmlElement | undefined {
  for (const child of el.elements ?? []) {
    if (child.type === "element" && child.name !== undefined && FILL_TAGS.has(child.name)) {
      return child;
    }
  }
  return undefined;
}

// ── Helper: read EG_ColorChoice directly from an element ──

function readDirectColor(el: XmlElement, ctx: ReadContext): SolidFillOptions {
  // Handles all six color kinds (srgbClr/schemeClr/hslClr/sysClr/prstClr/scrgbClr).
  const color = parseColorChoice(el, ctx);
  if (Object.keys(color).length > 0) return color;
  // Fallback: try solidFill wrapper (for backward compat with old-generated XML)
  const solidFill = findChild(el, "a:solidFill");
  if (solidFill) return parse(solidFillDesc, solidFill, ctx);
  return { value: "" };
}
