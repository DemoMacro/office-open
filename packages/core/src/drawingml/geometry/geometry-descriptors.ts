/**
 * Geometry descriptors for DrawingML shapes.
 *
 * @module
 */

import { escapeXml, findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type { CustomDescriptor } from "../../descriptor";
import { stringify, parse } from "../../descriptor";
import type { GeometryGuide } from "./adjustment-values";
import type {
  CustomGeometryOptions,
  PathOptions,
  PathCommand,
  AdjPoint,
  AdjustHandle,
  XYAdjustHandle,
  PolarAdjustHandle,
  ConnectionSite,
  GeomRect,
} from "./custom-geometry";
import type { PresetGeometryOptions } from "./preset-geometry";

// ── Adjustment values helper ──

const adjustmentValuesDesc: CustomDescriptor<readonly GeometryGuide[]> = {
  kind: "custom",
  stringify(guides, _ctx) {
    if (!guides || guides.length === 0) return "<a:avLst/>";
    const inner = guides
      .map((g) => `<a:gd name="${escapeXml(g.name)}" fmla="${escapeXml(g.formula)}"/>`)
      .join("");
    return `<a:avLst>${inner}</a:avLst>`;
  },
  parse(el, _ctx) {
    const result: Partial<GeometryGuide>[] = [];
    if (el.elements) {
      for (const child of el.elements) {
        if (child.name === "a:gd" && child.attributes) {
          const name = child.attributes["name"];
          const fmla = child.attributes["fmla"];
          if (name !== undefined && fmla !== undefined) {
            result.push({ name: String(name), formula: String(fmla) });
          }
        }
      }
    }
    return result as Partial<readonly GeometryGuide[]>;
  },
};

// ── Preset geometry descriptor ──

export const presetGeometryDesc: CustomDescriptor<PresetGeometryOptions> = {
  kind: "custom",
  stringify(opts, ctx) {
    const prst = opts.preset ?? "rect";
    let avXml = "";
    if (opts.adjustmentValues) {
      avXml = stringify(adjustmentValuesDesc, opts.adjustmentValues, ctx) ?? "<a:avLst/>";
    } else {
      avXml = "<a:avLst/>";
    }
    return `<a:prstGeom prst="${escapeXml(prst)}">${avXml}</a:prstGeom>`;
  },
  parse(el, ctx) {
    const result: Partial<PresetGeometryOptions> = {};
    if (el.attributes?.["prst"] !== undefined) {
      result.preset = String(el.attributes["prst"]);
    }
    const avLst = findChild(el, "a:avLst");
    if (avLst) {
      const guides = parse(adjustmentValuesDesc, avLst, ctx) as unknown;
      if (Array.isArray(guides) && guides.length > 0) {
        result.adjustmentValues = guides as GeometryGuide[];
      }
    }
    return result;
  },
};

// ── Path command helpers ──

function stringifyAdjPoint(pt: AdjPoint): string {
  return `<a:pt x="${escapeXml(pt.x)}" y="${escapeXml(pt.y)}"/>`;
}

function stringifyPathCommand(cmd: PathCommand): string {
  switch (cmd.command) {
    case "moveTo":
      return `<a:moveTo>${stringifyAdjPoint(cmd.point)}</a:moveTo>`;
    case "lineTo":
      return `<a:lnTo>${stringifyAdjPoint(cmd.point)}</a:lnTo>`;
    case "arcTo":
      return (
        `<a:arcTo wR="${escapeXml(cmd.widthRadius)}" hR="${escapeXml(cmd.heightRadius)}"` +
        ` stAng="${escapeXml(cmd.startAngle)}" swAng="${escapeXml(cmd.sweepAngle)}"/>`
      );
    case "quadBezTo": {
      const pts = cmd.points.map(stringifyAdjPoint).join("");
      return `<a:quadBezTo>${pts}</a:quadBezTo>`;
    }
    case "cubicBezTo": {
      const pts = cmd.points.map(stringifyAdjPoint).join("");
      return `<a:cubicBezTo>${pts}</a:cubicBezTo>`;
    }
    case "close":
      return "<a:close/>";
  }
}

function stringifyPath(path: PathOptions): string {
  const attrs: string[] = [];
  if (path.w !== undefined) attrs.push(`w="${path.w}"`);
  if (path.h !== undefined) attrs.push(`h="${path.h}"`);
  if (path.fill !== undefined) attrs.push(`fill="${escapeXml(path.fill)}"`);
  if (path.stroke !== undefined) attrs.push(`stroke="${path.stroke}"`);
  if (path.extrusionOk !== undefined) attrs.push(`extrusionOk="${path.extrusionOk}"`);
  const attrStr = attrs.length ? " " + attrs.join(" ") : "";
  const cmds = path.commands.map(stringifyPathCommand).join("");
  if (!cmds && !attrStr) return "<a:path/>";
  return `<a:path${attrStr}>${cmds}</a:path>`;
}

function readAdjPoint(el: XmlElement): AdjPoint | undefined {
  if (!el.attributes) return undefined;
  const x = el.attributes["x"];
  const y = el.attributes["y"];
  if (x === undefined || y === undefined) return undefined;
  return { x: String(x), y: String(y) };
}

function readPathCommand(tag: string, el: XmlElement): PathCommand | undefined {
  switch (tag) {
    case "a:moveTo": {
      const pt = el.elements?.find((c) => c.name === "a:pt");
      if (!pt) return undefined;
      const point = readAdjPoint(pt);
      if (!point) return undefined;
      return { command: "moveTo", point };
    }
    case "a:lnTo": {
      const pt = el.elements?.find((c) => c.name === "a:pt");
      if (!pt) return undefined;
      const point = readAdjPoint(pt);
      if (!point) return undefined;
      return { command: "lineTo", point };
    }
    case "a:arcTo": {
      const a = el.attributes;
      if (!a) return undefined;
      return {
        command: "arcTo",
        widthRadius: String(a["wR"] ?? ""),
        heightRadius: String(a["hR"] ?? ""),
        startAngle: String(a["stAng"] ?? ""),
        sweepAngle: String(a["swAng"] ?? ""),
      };
    }
    case "a:quadBezTo": {
      const points = (el.elements ?? [])
        .filter((c) => c.name === "a:pt")
        .map(readAdjPoint)
        .filter((p): p is AdjPoint => p !== undefined);
      if (points.length < 2) return undefined;
      return { command: "quadBezTo", points: [points[0], points[1]] };
    }
    case "a:cubicBezTo": {
      const points = (el.elements ?? [])
        .filter((c) => c.name === "a:pt")
        .map(readAdjPoint)
        .filter((p): p is AdjPoint => p !== undefined);
      if (points.length < 3) return undefined;
      return { command: "cubicBezTo", points: [points[0], points[1], points[2]] };
    }
    case "a:close":
      return { command: "close" };
    default:
      return undefined;
  }
}

function readPath(el: XmlElement): Partial<PathOptions> {
  const result: Partial<PathOptions> = {};
  if (el.attributes) {
    if (el.attributes["w"] !== undefined) result.w = Number(el.attributes["w"]);
    if (el.attributes["h"] !== undefined) result.h = Number(el.attributes["h"]);
    if (el.attributes["fill"] !== undefined)
      result.fill = String(el.attributes["fill"]) as PathOptions["fill"];
    if (el.attributes["stroke"] !== undefined)
      result.stroke = el.attributes["stroke"] !== "0" && el.attributes["stroke"] !== "false";
    if (el.attributes["extrusionOk"] !== undefined)
      result.extrusionOk =
        el.attributes["extrusionOk"] !== "0" && el.attributes["extrusionOk"] !== "false";
  }
  const commands: PathCommand[] = [];
  if (el.elements) {
    for (const child of el.elements) {
      if (child.name) {
        const cmd = readPathCommand(child.name, child);
        if (cmd) commands.push(cmd);
      }
    }
  }
  if (commands.length > 0) result.commands = commands;
  return result;
}

// ── Guide list helper (avLst / gdLst) ──

function stringifyGuideList(tag: string, guides: readonly GeometryGuide[]): string {
  if (!guides || guides.length === 0) return `<${tag}/>`;
  const inner = guides
    .map((g) => `<a:gd name="${escapeXml(g.name)}" fmla="${escapeXml(g.formula)}"/>`)
    .join("");
  return `<${tag}>${inner}</${tag}>`;
}

function readGuideList(el: XmlElement): GeometryGuide[] {
  const result: GeometryGuide[] = [];
  if (el.elements) {
    for (const child of el.elements) {
      if (child.name === "a:gd" && child.attributes) {
        const name = child.attributes["name"];
        const fmla = child.attributes["fmla"];
        if (name !== undefined && fmla !== undefined) {
          result.push({ name: String(name), formula: String(fmla) });
        }
      }
    }
  }
  return result;
}

// ── Adjust handle helpers ──

function stringifyAdjustHandlePos(pos: { x: string; y: string }): string {
  return `<a:pos x="${escapeXml(pos.x)}" y="${escapeXml(pos.y)}"/>`;
}

function stringifyXYAdjustHandle(h: XYAdjustHandle): string {
  const attrs: string[] = [];
  if (h.guideRefX !== undefined) attrs.push(`gdRefX="${escapeXml(h.guideRefX)}"`);
  if (h.minX !== undefined) attrs.push(`minX="${escapeXml(h.minX)}"`);
  if (h.maxX !== undefined) attrs.push(`maxX="${escapeXml(h.maxX)}"`);
  if (h.guideRefY !== undefined) attrs.push(`gdRefY="${escapeXml(h.guideRefY)}"`);
  if (h.minY !== undefined) attrs.push(`minY="${escapeXml(h.minY)}"`);
  if (h.maxY !== undefined) attrs.push(`maxY="${escapeXml(h.maxY)}"`);
  const attrStr = attrs.length ? " " + attrs.join(" ") : "";
  return `<a:ahXY${attrStr}>${stringifyAdjustHandlePos(h.position)}</a:ahXY>`;
}

function stringifyPolarAdjustHandle(h: PolarAdjustHandle): string {
  const attrs: string[] = [];
  if (h.guideRefRadius !== undefined) attrs.push(`gdRefR="${escapeXml(h.guideRefRadius)}"`);
  if (h.minRadius !== undefined) attrs.push(`minR="${escapeXml(h.minRadius)}"`);
  if (h.maxRadius !== undefined) attrs.push(`maxR="${escapeXml(h.maxRadius)}"`);
  if (h.guideRefAngle !== undefined) attrs.push(`gdRefAng="${escapeXml(h.guideRefAngle)}"`);
  if (h.minAngle !== undefined) attrs.push(`minAng="${escapeXml(h.minAngle)}"`);
  if (h.maxAngle !== undefined) attrs.push(`maxAng="${escapeXml(h.maxAngle)}"`);
  const attrStr = attrs.length ? " " + attrs.join(" ") : "";
  return `<a:ahPolar${attrStr}>${stringifyAdjustHandlePos(h.position)}</a:ahPolar>`;
}

function stringifyAdjustHandle(h: AdjustHandle): string {
  return h.type === "xy" ? stringifyXYAdjustHandle(h) : stringifyPolarAdjustHandle(h);
}

function readAdjustHandlePos(el: XmlElement): { x: string; y: string } | undefined {
  const pos = el.elements?.find((c) => c.name === "a:pos");
  if (!pos?.attributes) return undefined;
  const x = pos.attributes["x"];
  const y = pos.attributes["y"];
  if (x === undefined || y === undefined) return undefined;
  return { x: String(x), y: String(y) };
}

function readXYAdjustHandle(el: XmlElement): XYAdjustHandle | undefined {
  const position = readAdjustHandlePos(el);
  if (!position) return undefined;
  const result: XYAdjustHandle = { type: "xy", position };
  if (el.attributes?.["gdRefX"] !== undefined) result.guideRefX = String(el.attributes["gdRefX"]);
  if (el.attributes?.["minX"] !== undefined) result.minX = String(el.attributes["minX"]);
  if (el.attributes?.["maxX"] !== undefined) result.maxX = String(el.attributes["maxX"]);
  if (el.attributes?.["gdRefY"] !== undefined) result.guideRefY = String(el.attributes["gdRefY"]);
  if (el.attributes?.["minY"] !== undefined) result.minY = String(el.attributes["minY"]);
  if (el.attributes?.["maxY"] !== undefined) result.maxY = String(el.attributes["maxY"]);
  return result;
}

function readPolarAdjustHandle(el: XmlElement): PolarAdjustHandle | undefined {
  const position = readAdjustHandlePos(el);
  if (!position) return undefined;
  const result: PolarAdjustHandle = { type: "polar", position };
  if (el.attributes?.["gdRefR"] !== undefined)
    result.guideRefRadius = String(el.attributes["gdRefR"]);
  if (el.attributes?.["minR"] !== undefined) result.minRadius = String(el.attributes["minR"]);
  if (el.attributes?.["maxR"] !== undefined) result.maxRadius = String(el.attributes["maxR"]);
  if (el.attributes?.["gdRefAng"] !== undefined)
    result.guideRefAngle = String(el.attributes["gdRefAng"]);
  if (el.attributes?.["minAng"] !== undefined) result.minAngle = String(el.attributes["minAng"]);
  if (el.attributes?.["maxAng"] !== undefined) result.maxAngle = String(el.attributes["maxAng"]);
  return result;
}

// ── Connection site helpers ──

function stringifyConnectionSite(site: ConnectionSite): string {
  return `<a:cxn ang="${escapeXml(site.angle)}">${stringifyAdjustHandlePos(site.position)}</a:cxn>`;
}

function readConnectionSite(el: XmlElement): ConnectionSite | undefined {
  const ang = el.attributes?.["ang"];
  if (ang === undefined) return undefined;
  const position = readAdjustHandlePos(el);
  if (!position) return undefined;
  return { angle: String(ang), position };
}

// ── Geom rect helpers ──

function stringifyGeomRect(rect: GeomRect): string {
  return (
    `<a:rect l="${escapeXml(rect.left)}" t="${escapeXml(rect.top)}"` +
    ` r="${escapeXml(rect.right)}" b="${escapeXml(rect.bottom)}"/>`
  );
}

function readGeomRect(el: XmlElement): GeomRect | undefined {
  const a = el.attributes;
  if (!a) return undefined;
  const l = a["l"];
  const t = a["t"];
  const r = a["r"];
  const b = a["b"];
  if (l === undefined || t === undefined || r === undefined || b === undefined) return undefined;
  return { left: String(l), top: String(t), right: String(r), bottom: String(b) };
}

// ── Custom geometry descriptor ──

export const customGeometryDesc: CustomDescriptor<CustomGeometryOptions> = {
  kind: "custom",
  stringify(opts, _ctx) {
    const parts: string[] = [];

    // a:avLst
    if (opts.adjustmentValues && opts.adjustmentValues.length > 0) {
      parts.push(stringifyGuideList("a:avLst", opts.adjustmentValues));
    } else {
      parts.push("<a:avLst/>");
    }

    // a:gdLst
    if (opts.guides && opts.guides.length > 0) {
      parts.push(stringifyGuideList("a:gdLst", opts.guides));
    }

    // a:ahLst
    if (opts.adjustHandles && opts.adjustHandles.length > 0) {
      const inner = opts.adjustHandles.map(stringifyAdjustHandle).join("");
      parts.push(`<a:ahLst>${inner}</a:ahLst>`);
    }

    // a:cxnLst
    if (opts.connectionSites && opts.connectionSites.length > 0) {
      const inner = opts.connectionSites.map(stringifyConnectionSite).join("");
      parts.push(`<a:cxnLst>${inner}</a:cxnLst>`);
    }

    // a:rect
    if (opts.textRect) {
      parts.push(stringifyGeomRect(opts.textRect));
    }

    // a:pathLst (required)
    const pathsXml = opts.pathList.map(stringifyPath).join("");
    parts.push(`<a:pathLst>${pathsXml}</a:pathLst>`);

    return `<a:custGeom>${parts.join("")}</a:custGeom>`;
  },
  parse(el, _ctx) {
    const result: Partial<CustomGeometryOptions> = {};

    // a:avLst
    const avLst = findChild(el, "a:avLst");
    if (avLst) {
      const guides = readGuideList(avLst);
      if (guides.length > 0) result.adjustmentValues = guides;
    }

    // a:gdLst
    const gdLst = findChild(el, "a:gdLst");
    if (gdLst) {
      const guides = readGuideList(gdLst);
      if (guides.length > 0) result.guides = guides;
    }

    // a:ahLst
    const ahLst = findChild(el, "a:ahLst");
    if (ahLst?.elements) {
      const handles: AdjustHandle[] = [];
      for (const child of ahLst.elements) {
        if (child.name === "a:ahXY") {
          const h = readXYAdjustHandle(child);
          if (h) handles.push(h);
        } else if (child.name === "a:ahPolar") {
          const h = readPolarAdjustHandle(child);
          if (h) handles.push(h);
        }
      }
      if (handles.length > 0) result.adjustHandles = handles;
    }

    // a:cxnLst
    const cxnLst = findChild(el, "a:cxnLst");
    if (cxnLst?.elements) {
      const sites: ConnectionSite[] = [];
      for (const child of cxnLst.elements) {
        if (child.name === "a:cxn") {
          const site = readConnectionSite(child);
          if (site) sites.push(site);
        }
      }
      if (sites.length > 0) result.connectionSites = sites;
    }

    // a:rect
    const rect = findChild(el, "a:rect");
    if (rect) {
      const textRect = readGeomRect(rect);
      if (textRect) result.textRect = textRect;
    }

    // a:pathLst (required)
    const pathLst = findChild(el, "a:pathLst");
    if (pathLst?.elements) {
      const paths: PathOptions[] = [];
      for (const child of pathLst.elements) {
        if (child.name === "a:path") {
          const p = readPath(child);
          if (p.commands && p.commands.length > 0) {
            paths.push(p as PathOptions);
          }
        }
      }
      if (paths.length > 0) result.pathList = paths;
    }

    return result;
  },
};
