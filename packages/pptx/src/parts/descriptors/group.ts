/**
 * Group shape descriptor for PPTX.
 *
 * @module
 */

import { convertEmuToPixels, convertPixelsToEmu } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attrBool, attrNum, findChild } from "@office-open/xml";
import { escapeXml } from "@office-open/xml";
import type { SlideChild as LegacySlideChild } from "@parts/slide/slide-child";

import type { PptxWriteContext } from "../../context";
import { parseChild, stringifyChild } from "./bridge";

// ── Types ──

export interface GroupShapeDescriptorOptions {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  rotation?: number;
  flipHorizontal?: boolean;
  children?: LegacySlideChild[];
}

// ── ID counter ──

let _nextGroupId = 200;

// ── GroupShape (p:grpSp) descriptor ──

export const groupShapeDesc: CustomDescriptor<GroupShapeDescriptorOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const descCtx = ctx as unknown as PptxWriteContext;
    const id = _nextGroupId++;
    const name = "Group";

    const x = convertPixelsToEmu(opts.x ?? 0);
    const y = convertPixelsToEmu(opts.y ?? 0);
    const w = convertPixelsToEmu(opts.width ?? 100);
    const h = convertPixelsToEmu(opts.height ?? 100);

    const xfrmAttrs: string[] = [];
    if (opts.flipHorizontal) xfrmAttrs.push(' flipH="1"');
    if (opts.rotation !== undefined) xfrmAttrs.push(` rot="${opts.rotation}"`);

    const parts: string[] = [];

    // p:nvGrpSpPr
    parts.push(
      `<p:nvGrpSpPr><p:cNvPr id="${id}" name="${escapeXml(name)}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`,
    );

    // p:grpSpPr
    parts.push(
      `<p:grpSpPr><a:xfrm${xfrmAttrs.join("")}>` +
        `<a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/>` +
        `<a:chOff x="${x}" y="${y}"/><a:chExt cx="${w}" cy="${h}"/>` +
        `</a:xfrm></p:grpSpPr>`,
    );

    // Children
    if (opts.children) {
      for (const child of opts.children) {
        const xml = stringifyChild(child, descCtx);
        if (xml) parts.push(xml);
      }
    }

    return `<p:grpSp>${parts.join("")}</p:grpSp>`;
  },

  parse(el, _ctx) {
    const result: Record<string, unknown> = {};

    const grpSpPr = findChild(el, "p:grpSpPr");
    if (grpSpPr) {
      const xfrm = findChild(grpSpPr, "a:xfrm");
      if (xfrm) {
        const off = findChild(xfrm, "a:off");
        if (off) {
          const x = attrNum(off, "x");
          if (x !== undefined) result.x = convertEmuToPixels(x);
          const y = attrNum(off, "y");
          if (y !== undefined) result.y = convertEmuToPixels(y);
        }
        const ext = findChild(xfrm, "a:ext");
        if (ext) {
          const cx = attrNum(ext, "cx");
          if (cx !== undefined) result.width = convertEmuToPixels(cx);
          const cy = attrNum(ext, "cy");
          if (cy !== undefined) result.height = convertEmuToPixels(cy);
        }
        const rot = attrNum(xfrm, "rot");
        if (rot !== undefined) result.rotation = rot;
        if (attrBool(xfrm, "flipH")) result.flipHorizontal = true;
      }
    }

    // Children — recursive parse
    const groupChildren: LegacySlideChild[] = [];
    for (const child of el.elements ?? []) {
      if (child.name === "p:nvGrpSpPr" || child.name === "p:grpSpPr") continue;
      const parsed = parseChild(child, _ctx);
      if (parsed !== undefined) groupChildren.push(parsed);
    }
    if (groupChildren.length > 0) result.children = groupChildren;

    return result as Partial<GroupShapeDescriptorOptions>;
  },
};
