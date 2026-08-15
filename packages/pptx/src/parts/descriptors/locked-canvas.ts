/**
 * Locked Canvas descriptor for PPTX.
 *
 * Produces a p:graphicFrame with lc:lockedCanvas.
 *
 * @module
 */

import { convertToEmu } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import {
  stringifyNonVisualDrawingProperties,
  parseNonVisualDrawingProperties,
  shapePropertiesDesc,
  textBodyDesc,
} from "@office-open/core/drawing";
import { attrNum, findChild } from "@office-open/xml";

import type { LockedCanvasFrameOptions, LockedCanvasShapeOptions } from "../locked-canvas-frame";

// ── ID counters ──

let _nextLockedCanvasId = 2048;
let _nextCanvasShapeId = 2;

// ── Locked Canvas descriptor ──

export const lockedCanvasDesc: CustomDescriptor<LockedCanvasFrameOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const id = opts.id ?? _nextLockedCanvasId++;
    const name = opts.name ?? `Locked Canvas ${id}`;

    const x = convertToEmu(opts.x ?? 0);
    const y = convertToEmu(opts.y ?? 0);
    const cx = convertToEmu(opts.width ?? 0);
    const cy = convertToEmu(opts.height ?? 0);

    const parts: string[] = [];

    // p:nvGraphicFramePr
    parts.push(
      `<p:nvGraphicFramePr>${stringifyNonVisualDrawingProperties("p:cNvPr", id, opts, name)}` +
        `<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>` +
        `<p:nvPr/></p:nvGraphicFramePr>`,
    );

    // p:xfrm
    parts.push(`<p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>`);

    // a:graphic > a:graphicData > lc:lockedCanvas
    const grpSpPr =
      `<a:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="${cx}" cy="${cy}"/></a:xfrm></a:grpSpPr>`;
    const canvasChildren = buildCanvasChildren(opts.children, ctx);

    parts.push(
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas">` +
        `<lc:lockedCanvas xmlns:lc="http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas">` +
        `<a:nvGrpSpPr><a:cNvPr id="0" name=""/><a:cNvGrpSpPr/></a:nvGrpSpPr>` +
        `${grpSpPr}${canvasChildren}` +
        `</lc:lockedCanvas></a:graphicData></a:graphic>`,
    );

    return `<p:graphicFrame>${parts.join("")}</p:graphicFrame>`;
  },

  parse(el, _ctx) {
    const result: Partial<LockedCanvasFrameOptions> = {};

    // id, name from p:nvGraphicFramePr/p:cNvPr
    const nvGrFrm = findChild(el, "p:nvGraphicFramePr");
    if (nvGrFrm) {
      const cnvPr = findChild(nvGrFrm, "p:cNvPr");
      if (cnvPr) {
        Object.assign(result, parseNonVisualDrawingProperties(cnvPr));
        const id = attrNum(cnvPr, "id");
        if (id !== undefined) result.id = id;
      }
    }

    // x, y, width, height from p:xfrm
    const xfrm = findChild(el, "p:xfrm");
    if (xfrm) {
      const off = findChild(xfrm, "a:off");
      if (off) {
        const x = attrNum(off, "x");
        if (x !== undefined) result.x = x;
        const y = attrNum(off, "y");
        if (y !== undefined) result.y = y;
      }
      const ext = findChild(xfrm, "a:ext");
      if (ext) {
        const cx = attrNum(ext, "cx");
        if (cx !== undefined) result.width = cx;
        const cy = attrNum(ext, "cy");
        if (cy !== undefined) result.height = cy;
      }
    }

    // Navigate to a:graphic/a:graphicData/lc:lockedCanvas
    const graphic = findChild(el, "a:graphic");
    const graphicData = graphic ? findChild(graphic, "a:graphicData") : undefined;
    const lockedCanvas = graphicData ? findChild(graphicData, "lc:lockedCanvas") : undefined;

    if (lockedCanvas) {
      const children: LockedCanvasShapeOptions[] = [];
      for (const child of lockedCanvas.elements ?? []) {
        if (child.name !== "a:sp") continue;
        const shape: Partial<LockedCanvasShapeOptions> = {};

        // Full CT_ShapeProperties round-trip (geometry/fill/outline/effects/…)
        const spPr = findChild(child, "a:spPr");
        if (spPr) {
          Object.assign(shape, shapePropertiesDesc.parse(spPr, _ctx));
        }

        // textBody from a:txSp/a:txBody/a:p/a:r/a:t
        const txSp = findChild(child, "a:txSp");
        if (txSp) {
          const txBody = findChild(txSp, "a:txBody");
          if (txBody) {
            const txt = collectRunText(txBody);
            if (txt) shape.textBody = txt;
          }
        }

        children.push(shape as LockedCanvasShapeOptions);
      }
      if (children.length > 0) result.children = children;
    }

    return result as LockedCanvasFrameOptions;
  },
};

function buildCanvasChildren(
  children: LockedCanvasShapeOptions[] | undefined,
  ctx: Parameters<typeof shapePropertiesDesc.stringify>[1],
): string {
  if (!children || children.length === 0) return "";

  const parts: string[] = [];
  for (const { textBody, ...spPr } of children) {
    const id = _nextCanvasShapeId++;
    const spPrContent = shapePropertiesDesc.stringify(
      { ...spPr, geometry: spPr.geometry ?? "rect" },
      ctx,
    );

    const txBodyContent = textBodyDesc.stringify(
      textBody !== undefined ? { text: textBody } : {},
      ctx,
    );

    parts.push(
      `<a:sp>` +
        `<a:nvSpPr><a:cNvPr id="${id}" name="Shape ${id}"/><a:cNvSpPr/></a:nvSpPr>` +
        `<a:spPr>${spPrContent ?? ""}</a:spPr>` +
        `<a:txSp><a:txBody>${txBodyContent ?? ""}</a:txBody><a:useSpRect/></a:txSp>` +
        `</a:sp>`,
    );
  }

  return parts.join("");
}

/** Collect text from a:r/a:t children within a txBody element. */
function collectRunText(el: {
  elements?: Array<{ name?: string; text?: string | number | boolean; elements?: Array<unknown> }>;
}): string {
  const parts: string[] = [];
  for (const child of el.elements ?? []) {
    if (child.name === "a:r") {
      // Look for a:t child
      for (const sub of (child.elements ?? []) as Array<{
        name?: string;
        text?: string | number | boolean;
      }>) {
        if (sub.name === "a:t" && sub.text !== undefined) {
          parts.push(String(sub.text));
        }
      }
    } else if (child.name === "a:p") {
      parts.push(collectRunText(child as Parameters<typeof collectRunText>[0]));
    }
  }
  return parts.join("");
}
