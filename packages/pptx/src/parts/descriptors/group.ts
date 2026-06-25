/**
 * Group shape descriptor for PPTX.
 *
 * @module
 */

import { convertEmuToPixels, convertPixelsToEmu } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { parse, stringify } from "@office-open/core/descriptor";
import { effectListDesc, fillDesc } from "@office-open/core/drawingml";
import type { FillOptions as CoreFillOptions } from "@office-open/core/drawingml";
import { attrBool, attrNum, findChild } from "@office-open/xml";
import { escapeXml } from "@office-open/xml";
import type { SlideChild as LegacySlideChild } from "@parts/slide/slide-child";
import { toEffectListOptions, type EffectsOptions } from "@shared/drawingml/effects";

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
  /** Child coordinate system offset (a:chOff). Defaults to {x,y} when omitted. */
  childOffset?: { x: number; y: number };
  /** Child coordinate system extent (a:chExt). Defaults to {width,height} when omitted. */
  childExtent?: { cx: number; cy: number };
  /** Group-level fill (EG_FillProperties on grpSpPr). */
  fill?: CoreFillOptions;
  /** Group-level effects (EG_EffectProperties on grpSpPr). */
  effects?: EffectsOptions;
  children?: LegacySlideChild[];
}

// ── ID counter ──

let _nextGroupId = 200;

// ── GroupShape (p:grpSp) descriptor ──

export const groupShapeDesc: CustomDescriptor<GroupShapeDescriptorOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const descCtx = ctx as PptxWriteContext;
    const id = _nextGroupId++;
    const name = "Group";

    const x = convertPixelsToEmu(opts.x ?? 0);
    const y = convertPixelsToEmu(opts.y ?? 0);
    const w = convertPixelsToEmu(opts.width ?? 100);
    const h = convertPixelsToEmu(opts.height ?? 100);

    // chOff/chExt default to off/ext when the child coordinate system is unchanged.
    const chOffX = opts.childOffset ? convertPixelsToEmu(opts.childOffset.x) : x;
    const chOffY = opts.childOffset ? convertPixelsToEmu(opts.childOffset.y) : y;
    const chExtCx = opts.childExtent ? convertPixelsToEmu(opts.childExtent.cx) : w;
    const chExtCy = opts.childExtent ? convertPixelsToEmu(opts.childExtent.cy) : h;

    const xfrmAttrs: string[] = [];
    if (opts.flipHorizontal) xfrmAttrs.push(' flipH="1"');
    if (opts.rotation !== undefined) xfrmAttrs.push(` rot="${opts.rotation}"`);

    const grpSpPrParts: string[] = [
      `<a:xfrm${xfrmAttrs.join("")}>` +
        `<a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/>` +
        `<a:chOff x="${chOffX}" y="${chOffY}"/><a:chExt cx="${chExtCx}" cy="${chExtCy}"/>` +
        `</a:xfrm>`,
    ];
    if (opts.fill) {
      const fillXml = stringify(fillDesc, opts.fill, descCtx);
      if (fillXml) grpSpPrParts.push(fillXml);
    }
    if (opts.effects) {
      const effectListOpts = toEffectListOptions(opts.effects);
      if (effectListOpts) {
        const fxXml = stringify(effectListDesc, effectListOpts, descCtx);
        if (fxXml) grpSpPrParts.push(fxXml);
      }
    }

    const parts: string[] = [];

    // p:nvGrpSpPr
    parts.push(
      `<p:nvGrpSpPr><p:cNvPr id="${id}" name="${escapeXml(name)}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`,
    );

    // p:grpSpPr
    parts.push(`<p:grpSpPr>${grpSpPrParts.join("")}</p:grpSpPr>`);

    // Children
    if (opts.children) {
      for (const child of opts.children) {
        const xml = stringifyChild(child, descCtx);
        if (xml) parts.push(xml);
      }
    }

    return `<p:grpSp>${parts.join("")}</p:grpSp>`;
  },

  parse(el, ctx) {
    const result: Partial<GroupShapeDescriptorOptions> = {};

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
        const chOff = findChild(xfrm, "a:chOff");
        if (chOff) {
          const cox = attrNum(chOff, "x");
          const coy = attrNum(chOff, "y");
          if (cox !== undefined && coy !== undefined) {
            result.childOffset = { x: convertEmuToPixels(cox), y: convertEmuToPixels(coy) };
          }
        }
        const chExt = findChild(xfrm, "a:chExt");
        if (chExt) {
          const ccx = attrNum(chExt, "cx");
          const ccy = attrNum(chExt, "cy");
          if (ccx !== undefined && ccy !== undefined) {
            result.childExtent = { cx: convertEmuToPixels(ccx), cy: convertEmuToPixels(ccy) };
          }
        }
        const rot = attrNum(xfrm, "rot");
        if (rot !== undefined) result.rotation = rot;
        if (attrBool(xfrm, "flipH")) result.flipHorizontal = true;
      }
      const fillResult = parse(fillDesc, grpSpPr, ctx);
      if (fillResult && Object.keys(fillResult).length > 0) {
        result.fill = fillResult as GroupShapeDescriptorOptions["fill"];
      }
      const effectLst = findChild(grpSpPr, "a:effectLst");
      if (effectLst) {
        result.effects = parse(
          effectListDesc,
          effectLst,
          ctx,
        ) as GroupShapeDescriptorOptions["effects"];
      }
    }

    // Children — recursive parse
    const groupChildren: LegacySlideChild[] = [];
    for (const child of el.elements ?? []) {
      if (child.name === "p:nvGrpSpPr" || child.name === "p:grpSpPr") continue;
      const parsed = parseChild(child, ctx);
      if (parsed !== undefined) groupChildren.push(parsed);
    }
    if (groupChildren.length > 0) result.children = groupChildren;

    return result as GroupShapeDescriptorOptions;
  },
};
