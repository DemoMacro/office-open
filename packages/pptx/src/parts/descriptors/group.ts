/**
 * Group shape descriptor for PPTX.
 *
 * @module
 */

import { convertToEmu } from "@office-open/core";
import type { UniversalMeasure } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import {
  groupShapePropertiesDesc,
  parseNonVisualDrawingProperties,
  stringifyNonVisualDrawingProperties,
} from "@office-open/core/drawingml";
import type {
  BlackWhiteMode,
  EffectListOptions,
  FillOptions as CoreFillOptions,
  NonVisualDrawingPropertiesOptions,
} from "@office-open/core/drawingml";
import { attrNum, findChild } from "@office-open/xml";
import type { SlideChild as LegacySlideChild } from "@parts/slide/slide-child";

import type { PptxWriteContext } from "../../context";
import { parseChild, stringifyChild } from "./bridge";

// ── Types ──

export interface GroupShapeDescriptorOptions extends NonVisualDrawingPropertiesOptions {
  id?: number;
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  /** Rotation angle in degrees (e.g., 45 = 45°). */
  rotation?: number;
  flipHorizontal?: boolean;
  /** Child coordinate system offset (a:chOff). Defaults to {x,y} when omitted. */
  childOffset?: { x: number | UniversalMeasure; y: number | UniversalMeasure };
  /** Child coordinate system extent (a:chExt). Defaults to {width,height} when omitted. */
  childExtent?: { cx: number | UniversalMeasure; cy: number | UniversalMeasure };
  /** Group-level fill (EG_FillProperties on grpSpPr). */
  fill?: CoreFillOptions;
  /** Group-level effects (EG_EffectProperties on grpSpPr). */
  effects?: EffectListOptions;
  /** @bwMode container attribute (ST_BlackWhiteMode) on p:grpSpPr. */
  bwMode?: BlackWhiteMode;
  children?: LegacySlideChild[];
}

// ── ID counter ──

let _nextGroupId = 200;

// ── GroupShape (p:grpSp) descriptor ──

export const groupShapeDesc: CustomDescriptor<GroupShapeDescriptorOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const descCtx = ctx as PptxWriteContext;
    const id = opts.id ?? _nextGroupId++;
    const name = opts.name ?? "Group";

    const x = convertToEmu(opts.x ?? 0);
    const y = convertToEmu(opts.y ?? 0);
    const w = convertToEmu(opts.width ?? "100px");
    const h = convertToEmu(opts.height ?? "100px");

    const grpSpPrContent =
      groupShapePropertiesDesc.stringify(
        {
          x,
          y,
          width: w,
          height: h,
          ...(opts.flipHorizontal !== undefined ? { flipHorizontal: opts.flipHorizontal } : {}),
          ...(opts.rotation !== undefined ? { rotation: opts.rotation } : {}),
          // chOff/chExt default to off/ext when the child coordinate system is unchanged.
          childOffsetX: opts.childOffset ? convertToEmu(opts.childOffset.x) : x,
          childOffsetY: opts.childOffset ? convertToEmu(opts.childOffset.y) : y,
          childExtentWidth: opts.childExtent ? convertToEmu(opts.childExtent.cx) : w,
          childExtentHeight: opts.childExtent ? convertToEmu(opts.childExtent.cy) : h,
          ...(opts.fill !== undefined ? { fill: opts.fill } : {}),
          ...(opts.effects !== undefined ? { effects: opts.effects } : {}),
        },
        descCtx,
      ) ?? "";

    const grpSpPrAttrs = opts.bwMode ? ` bwMode="${opts.bwMode}"` : "";

    const parts: string[] = [];

    // p:nvGrpSpPr
    parts.push(
      `<p:nvGrpSpPr>${stringifyNonVisualDrawingProperties("p:cNvPr", id, opts, name)}<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`,
    );

    // p:grpSpPr
    parts.push(`<p:grpSpPr${grpSpPrAttrs}>${grpSpPrContent}</p:grpSpPr>`);

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

    // id + name from p:nvGrpSpPr/p:cNvPr
    const nvGrpSpPr = findChild(el, "p:nvGrpSpPr");
    if (nvGrpSpPr) {
      const cNvPr = findChild(nvGrpSpPr, "p:cNvPr");
      if (cNvPr) {
        Object.assign(result, parseNonVisualDrawingProperties(cNvPr));
        const id = attrNum(cNvPr, "id");
        if (id !== undefined) result.id = id;
      }
    }

    const grpSpPr = findChild(el, "p:grpSpPr");
    if (grpSpPr) {
      const props = groupShapePropertiesDesc.parse(grpSpPr, ctx);
      if (props.x !== undefined) result.x = props.x;
      if (props.y !== undefined) result.y = props.y;
      if (props.width !== undefined) result.width = props.width;
      if (props.height !== undefined) result.height = props.height;
      if (props.rotation !== undefined) result.rotation = props.rotation;
      if (props.flipHorizontal) result.flipHorizontal = true;
      if (props.childOffsetX !== undefined && props.childOffsetY !== undefined) {
        result.childOffset = { x: props.childOffsetX, y: props.childOffsetY };
      }
      if (props.childExtentWidth !== undefined && props.childExtentHeight !== undefined) {
        result.childExtent = { cx: props.childExtentWidth, cy: props.childExtentHeight };
      }
      if (props.fill !== undefined) result.fill = props.fill;
      if (props.effects !== undefined) result.effects = props.effects;
      const bwMode = grpSpPr.attributes?.["bwMode"];
      if (bwMode !== undefined) result.bwMode = bwMode as BlackWhiteMode;
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
