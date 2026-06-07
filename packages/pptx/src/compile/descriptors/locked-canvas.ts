/**
 * Locked Canvas descriptor for PPTX.
 *
 * Produces a p:graphicFrame with lc:lockedCanvas.
 *
 * @module
 */

import { convertPixelsToEmu } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { escapeXml } from "@office-open/xml";

// ── Types ──

export interface LockedCanvasShapeDescriptorOptions {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  geometry?: string;
  fill?: string;
  textBody?: string;
}

export interface LockedCanvasDescriptorOptions {
  id?: number;
  name?: string;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  children?: readonly LockedCanvasShapeDescriptorOptions[];
}

// ── ID counters ──

let _nextLockedCanvasId = 2048;
let _nextCanvasShapeId = 2;

// ── Locked Canvas descriptor ──

export const lockedCanvasDesc: CustomDescriptor<LockedCanvasDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const id = opts.id ?? _nextLockedCanvasId++;
    const name = opts.name ?? `Locked Canvas ${id}`;

    const x = convertPixelsToEmu(opts.x ?? 0);
    const y = convertPixelsToEmu(opts.y ?? 0);
    const cx = convertPixelsToEmu(opts.width ?? 0);
    const cy = convertPixelsToEmu(opts.height ?? 0);

    const parts: string[] = [];

    // p:nvGraphicFramePr
    parts.push(
      `<p:nvGraphicFramePr><p:cNvPr id="${id}" name="${escapeXml(name)}"/>` +
        `<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>` +
        `<p:nvPr/></p:nvGraphicFramePr>`,
    );

    // p:xfrm
    parts.push(`<p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>`);

    // a:graphic > a:graphicData > lc:lockedCanvas
    const grpSpPr =
      `<a:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="${cx}" cy="${cy}"/></a:xfrm></a:grpSpPr>`;
    const canvasChildren = buildCanvasChildren(opts.children);

    parts.push(
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas">` +
        `<lc:lockedCanvas xmlns:lc="http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas">` +
        `<a:nvGrpSpPr><a:cNvPr id="0" name=""/><a:cNvGrpSpPr/></a:nvGrpSpPr>` +
        `${grpSpPr}${canvasChildren}` +
        `</lc:lockedCanvas></a:graphicData></a:graphic>`,
    );

    return `<p:graphicFrame>${parts.join("")}</p:graphicFrame>`;
  },

  parse(_el, _ctx) {
    return {};
  },
};

function buildCanvasChildren(
  children: readonly LockedCanvasShapeDescriptorOptions[] | undefined,
): string {
  if (!children || children.length === 0) return "";

  const parts: string[] = [];
  for (const opts of children) {
    const id = _nextCanvasShapeId++;
    const spParts: string[] = [];

    // a:nvSpPr
    spParts.push(`<a:nvSpPr><a:cNvPr id="${id}" name="Shape ${id}"/><a:cNvSpPr/></a:nvSpPr>`);

    // a:spPr
    const spPrParts: string[] = [];
    const sx = convertPixelsToEmu(opts.x ?? 0);
    const sy = convertPixelsToEmu(opts.y ?? 0);
    const scx = convertPixelsToEmu(opts.width ?? 0);
    const scy = convertPixelsToEmu(opts.height ?? 0);
    spPrParts.push(`<a:xfrm><a:off x="${sx}" y="${sy}"/><a:ext cx="${scx}" cy="${scy}"/></a:xfrm>`);
    spPrParts.push(`<a:prstGeom prst="${opts.geometry ?? "rect"}"><a:avLst/></a:prstGeom>`);
    if (opts.fill) {
      spPrParts.push(`<a:solidFill><a:srgbClr val="${opts.fill}"/></a:solidFill>`);
    }
    spParts.push(`<a:spPr>${spPrParts.join("")}</a:spPr>`);

    // a:txSp > a:txBody > a:useSpRect
    spParts.push(
      `<a:txSp><a:txBody><a:bodyPr/><a:lstStyle/>` +
        `<a:p>${opts.textBody ? `<a:r><a:t>${escapeXml(opts.textBody)}</a:t></a:r>` : ""}</a:p>` +
        `</a:txBody><a:useSpRect/></a:txSp>`,
    );

    parts.push(`<a:sp>${spParts.join("")}</a:sp>`);
  }

  return parts.join("");
}
