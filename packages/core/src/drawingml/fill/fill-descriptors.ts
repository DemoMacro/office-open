/**
 * Fill descriptors for DrawingML EG_FillProperties.
 *
 * @module
 */

import { escapeXml } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { findChild } from "@office-open/xml";

import type { CustomDescriptor, ReadContext } from "../../descriptor";
import { stringify, parse } from "../../descriptor";
import { xsdPattern } from "../../util/mappings";
import { solidFillDesc, getColorDescriptor, parseColorChoice } from "../color/color-descriptors";
import type { SolidFillOptions } from "../color/solid-fill";
import type { FillOptions } from "./fill-options";
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
  const result: Partial<RelativeRect> = {};
  if (el.attributes?.["l"]) result.left = String(el.attributes["l"]);
  if (el.attributes?.["t"]) result.top = String(el.attributes["t"]);
  if (el.attributes?.["r"]) result.right = String(el.attributes["r"]);
  if (el.attributes?.["b"]) result.bottom = String(el.attributes["b"]);
  return result as RelativeRect;
}

function stringifyShade(shade: GradientShadeOptions): string {
  if ("angle" in shade) {
    const parts: string[] = [];
    if (shade.angle !== undefined) parts.push(`ang="${shade.angle}"`);
    if (shade.scaled !== undefined) parts.push(`scaled="${shade.scaled ? 1 : 0}"`);
    const attrStr = parts.length ? " " + parts.join(" ") : "";
    return `<a:lin${attrStr}/>`;
  }
  const pathShade = shade as PathShadeOptions;
  const parts: string[] = [];
  if (pathShade.path) parts.push(`path="${escapeXml(pathShade.path)}"`);
  const attrStr = parts.length ? " " + parts.join(" ") : "";
  if (pathShade.fillToRect) {
    return `<a:path${attrStr}>${stringifyRelativeRect("a:fillToRect", pathShade.fillToRect)}</a:path>`;
  }
  return `<a:path${attrStr}/>`;
}

// ── GradientFill descriptor ──

export const gradientFillDesc: CustomDescriptor<GradientFillOptions> = {
  kind: "custom",
  stringify(opts, ctx) {
    const parts: string[] = [];

    // Gradient stop list — a:gs expects EG_ColorChoice (direct color), NOT solidFill
    const stopsXml = opts.stops
      .map((stop) => {
        const desc = getColorDescriptor(stop.color);
        const colorXml = stringify(desc, stop.color, ctx);
        if (!colorXml) return `<a:gs pos="${stop.position}"/>`;
        return `<a:gs pos="${stop.position}">${colorXml}</a:gs>`;
      })
      .join("");
    parts.push(`<a:gsLst>${stopsXml}</a:gsLst>`);

    // Shade
    if (opts.shade) parts.push(stringifyShade(opts.shade));

    // Tile rect
    if (opts.tileRect) parts.push(stringifyRelativeRect("a:tileRect", opts.tileRect));

    // Attributes
    const attrParts: string[] = [];
    if (opts.flip) attrParts.push(`flip="${escapeXml(opts.flip)}"`);
    if (opts.rotateWithShape !== undefined)
      attrParts.push(`rotWithShape="${opts.rotateWithShape ? 1 : 0}"`);
    const attrStr = attrParts.length ? " " + attrParts.join(" ") : "";

    return `<a:gradFill${attrStr}>${parts.join("")}</a:gradFill>`;
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
          return { position: pos, color };
        });
    }

    // Shade (a:lin or a:path)
    const lin = findChild(el, "a:lin");
    if (lin) {
      const shade: Partial<LinearShadeOptions> = {};
      if (lin.attributes?.["ang"] !== undefined) shade.angle = Number(lin.attributes["ang"]);
      if (lin.attributes?.["scaled"] !== undefined) shade.scaled = lin.attributes["scaled"] !== "0";
      result.shade = shade as GradientShadeOptions;
    } else {
      const path = findChild(el, "a:path");
      if (path) {
        const shade: Partial<PathShadeOptions> = {};
        if (path.attributes?.["path"] !== undefined)
          shade.path = String(path.attributes["path"]) as PathShadeOptions["path"];
        const fillToRect = findChild(path, "a:fillToRect");
        if (fillToRect) shade.fillToRect = readRelativeRect(fillToRect);
        result.shade = shade as GradientShadeOptions;
      }
    }

    // Flip
    if (el.attributes?.["flip"] !== undefined)
      result.flip = String(el.attributes["flip"]) as GradientFillOptions["flip"];

    // Rotate with shape
    if (el.attributes?.["rotWithShape"] !== undefined)
      result.rotateWithShape = el.attributes["rotWithShape"] !== "0";

    // Tile rect
    const tileRect = findChild(el, "a:tileRect");
    if (tileRect) result.tileRect = readRelativeRect(tileRect);

    return result as GradientFillOptions;
  },
};

// ── PatternFill descriptor ──

export const patternFillDesc: CustomDescriptor<PatternFillOptions> = {
  kind: "custom",
  stringify(opts, ctx) {
    const parts: string[] = [];
    const prst = xsdPattern.to(opts.pattern);

    // a:fgClr/a:bgClr expect EG_ColorChoice (direct color), NOT solidFill
    if (opts.foregroundColor) {
      const desc = getColorDescriptor(opts.foregroundColor);
      const colorXml = stringify(desc, opts.foregroundColor, ctx);
      if (colorXml) parts.push(`<a:fgClr>${colorXml}</a:fgClr>`);
    }
    if (opts.backgroundColor) {
      const desc = getColorDescriptor(opts.backgroundColor);
      const colorXml = stringify(desc, opts.backgroundColor, ctx);
      if (colorXml) parts.push(`<a:bgClr>${colorXml}</a:bgClr>`);
    }

    const inner = parts.join("");
    return `<a:pattFill prst="${escapeXml(prst)}">${inner}</a:pattFill>`;
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

export const fillDesc: CustomDescriptor<FillOptions> = {
  kind: "custom",
  stringify(opts, ctx) {
    // String shorthand → solid fill
    if (typeof opts === "string") {
      return stringify(solidFillDesc, { value: opts.replace("#", "") } as SolidFillOptions, ctx);
    }

    switch (opts.type) {
      case "none":
        return "<a:noFill/>";

      case "solid": {
        const color =
          typeof opts.color === "string"
            ? ({ value: opts.color.replace("#", "") } as SolidFillOptions)
            : opts.color;
        return stringify(solidFillDesc, color, ctx);
      }

      case "gradient": {
        // Core API variant
        if ("options" in opts) {
          return stringify(gradientFillDesc, opts.options, ctx);
        }
        // Simplified API variant
        const gradOpts: GradientFillOptions = {
          stops: opts.stops.map((stop) => ({
            position: stop.position * 1000,
            color:
              typeof stop.color === "string"
                ? ({ value: stop.color.replace("#", "") } as SolidFillOptions)
                : stop.color,
          })),
        };
        if (!opts.path && opts.angle !== undefined) {
          gradOpts.shade = { angle: opts.angle * 60000, scaled: opts.scaled ?? true };
        }
        if (opts.path) {
          gradOpts.shade = { path: opts.path };
        }
        return stringify(gradientFillDesc, gradOpts, ctx);
      }

      case "blip":
        // Blip fills are handled in the blip module; skip here
        return undefined;

      case "pattern": {
        const patternOpts: PatternFillOptions = {
          pattern: opts.pattern as PatternFillOptions["pattern"],
          ...(opts.foregroundColor && {
            foregroundColor:
              typeof opts.foregroundColor === "string"
                ? ({ value: opts.foregroundColor.replace("#", "") } as SolidFillOptions)
                : opts.foregroundColor,
          }),
          ...(opts.backgroundColor && {
            backgroundColor:
              typeof opts.backgroundColor === "string"
                ? ({ value: opts.backgroundColor.replace("#", "") } as SolidFillOptions)
                : opts.backgroundColor,
          }),
        };
        return stringify(patternFillDesc, patternOpts, ctx);
      }

      case "group":
        return "<a:grpFill/>";
    }
  },
  parse(el, ctx) {
    // Resolve fill element — either el itself or a child
    const resolve = (tag: string): XmlElement | undefined =>
      el.name === tag ? el : findChild(el, tag);

    const noFill = resolve("a:noFill");
    if (noFill) {
      return { type: "none" } as FillOptions;
    }

    const solidFill = resolve("a:solidFill");
    if (solidFill) {
      const color = parse(solidFillDesc, solidFill, ctx) as SolidFillOptions;
      return { type: "solid", color } as FillOptions;
    }

    const gradFill = resolve("a:gradFill");
    if (gradFill) {
      const gradOpts = parse(gradientFillDesc, gradFill, ctx);
      return { type: "gradient", options: gradOpts as GradientFillOptions } as FillOptions;
    }

    const pattFill = resolve("a:pattFill");
    if (pattFill) {
      const pattOpts = parse(patternFillDesc, pattFill, ctx);
      return { type: "pattern", ...(pattOpts as PatternFillOptions) } as FillOptions;
    }

    const grpFill = resolve("a:grpFill");
    if (grpFill) {
      return { type: "group" } as FillOptions;
    }

    return { type: "none" } as FillOptions;
  },
};

// ── Helper: read EG_ColorChoice directly from an element ──

function readDirectColor(el: XmlElement, ctx: ReadContext): SolidFillOptions {
  // Handles all six color kinds (srgbClr/schemeClr/hslClr/sysClr/prstClr/scrgbClr).
  const color = parseColorChoice(el, ctx);
  if (Object.keys(color).length > 0) return color;
  // Fallback: try solidFill wrapper (for backward compat with old-generated XML)
  const solidFill = findChild(el, "a:solidFill");
  if (solidFill) return parse(solidFillDesc, solidFill, ctx) as SolidFillOptions;
  return { value: "" } as SolidFillOptions;
}
