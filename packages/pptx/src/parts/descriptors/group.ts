/**
 * Group shape descriptor for PPTX.
 *
 * @module
 */

import { convertToEmu } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import {
  groupLockingDesc,
  groupShapePropertiesDesc,
  stringifyNonVisualDrawingProperties,
} from "@office-open/core/drawing";
import type { BlackWhiteMode } from "@office-open/core/drawing";
import { findChild } from "@office-open/xml";
import type { SlideChild as LegacySlideChild } from "@parts/slide/slide-child";
import type { GroupOptions } from "@shared/shape/group-shape";

import type { PptxWriteContext } from "../../context";
import { parseChild, stringifyChild } from "./bridge";
import { readCnvPr } from "./shape";

// ── ID counter ──

let _nextGroupId = 200;

// ── GroupShape (p:grpSp) descriptor ──

export const groupShapeDesc: CustomDescriptor<GroupOptions> = {
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
          ...(opts.flipVertical !== undefined ? { flipVertical: opts.flipVertical } : {}),
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

    const grpSpPrAttrs = opts.blackWhiteMode ? ` bwMode="${opts.blackWhiteMode}"` : "";

    const parts: string[] = [];

    // p:nvGrpSpPr
    // null = explicit empty locks element; an object always emits the flags.
    const grpLocks =
      opts.locking === null
        ? "<a:grpSpLocks/>"
        : opts.locking
          ? (groupLockingDesc.stringify(opts.locking, descCtx) ?? "")
          : "";
    const cNvGrpSpPr = grpLocks ? `<p:cNvGrpSpPr>${grpLocks}</p:cNvGrpSpPr>` : "<p:cNvGrpSpPr/>";
    parts.push(
      `<p:nvGrpSpPr>${stringifyNonVisualDrawingProperties("p:cNvPr", id, opts, name)}${cNvGrpSpPr}<p:nvPr/></p:nvGrpSpPr>`,
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
    const result: Partial<GroupOptions> = {};

    // id + name from p:nvGrpSpPr/p:cNvPr
    Object.assign(result, readCnvPr(el, "p:nvGrpSpPr"));
    const nvGrpSpPr = findChild(el, "p:nvGrpSpPr");
    const grpLocks = nvGrpSpPr ? findChild(nvGrpSpPr, "p:cNvGrpSpPr") : undefined;
    const lockEl = grpLocks ? findChild(grpLocks, "a:grpSpLocks") : undefined;
    if (lockEl) {
      const locks = groupLockingDesc.parse(lockEl, ctx);
      result.locking = locks && Object.keys(locks).length > 0 ? locks : null;
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
      if (props.flipVertical) result.flipVertical = true;
      if (props.childOffsetX !== undefined && props.childOffsetY !== undefined) {
        result.childOffset = { x: props.childOffsetX, y: props.childOffsetY };
      }
      if (props.childExtentWidth !== undefined && props.childExtentHeight !== undefined) {
        result.childExtent = { cx: props.childExtentWidth, cy: props.childExtentHeight };
      }
      if (props.fill !== undefined) result.fill = props.fill;
      if (props.effects !== undefined) result.effects = props.effects;
      const bwMode = grpSpPr.attributes?.["bwMode"];
      if (bwMode !== undefined) result.blackWhiteMode = bwMode as BlackWhiteMode;
    }

    // Children — recursive parse
    const groupChildren: LegacySlideChild[] = [];
    for (const child of el.elements ?? []) {
      if (child.name === "p:nvGrpSpPr" || child.name === "p:grpSpPr") continue;
      const parsed = parseChild(child, ctx);
      if (parsed !== undefined) groupChildren.push(parsed);
    }
    if (groupChildren.length > 0) result.children = groupChildren;

    return result as GroupOptions;
  },
};
