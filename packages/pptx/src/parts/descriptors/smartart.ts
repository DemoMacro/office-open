/**
 * SmartArt (p:graphicFrame with diagram) descriptor for PPTX.
 *
 * Produces a graphicFrame with SmartArt relationship placeholders.
 * The actual diagram data is registered in PptxWriteContext for
 * separate compilation by the compiler.
 *
 * @module
 */

import { convertToEmu } from "@office-open/core";
import type { UniversalMeasure } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { createDataModel, type TreeNode } from "@office-open/core/smartart";
import { attr, findChild, findDeep } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { escapeXml } from "@office-open/xml";

import type { PptxWriteContext } from "../../context";
import { readPositionFromXfrm } from "./shape";

// ── Types ──

export interface SmartArtDescriptorOptions {
  id?: number;
  name?: string;
  x?: number | UniversalMeasure;
  y?: number | UniversalMeasure;
  width?: number | UniversalMeasure;
  height?: number | UniversalMeasure;
  /** Pre-generated SmartArt key (e.g. "smartart_1024"). If omitted, auto-generated. */
  smartArtKey?: string;
  /** Tree nodes for the diagram content. */
  nodes?: TreeNode[];
  /** Layout ID (e.g. "default", "process1", "hierarchy1"). */
  layout?: string;
  /** Quick style ID (e.g. "simple1", "moderate1"). */
  style?: string;
  /** Color transform ID (e.g. "accent1_2", "colorful1"). */
  color?: string;
}

// ── ID counter ──

let _nextSmartArtId = 1024;

// ── SmartArt descriptor ──

export const smartArtDesc: CustomDescriptor<SmartArtDescriptorOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const pptxCtx = ctx as PptxWriteContext;
    const id = opts.id ?? _nextSmartArtId++;
    const name = opts.name ?? `Diagram ${id}`;
    const saKey = opts.smartArtKey ?? pptxCtx.nextSmartArtKey();

    const layoutId = opts.layout ?? "default";
    const styleId = opts.style ?? "simple1";
    const colorId = opts.color ?? "accent1_2";

    // Register SmartArt data with context
    if (opts.nodes && opts.nodes.length > 0) {
      const body = createDataModel(opts.nodes, layoutId, styleId, colorId);
      const dataModelXml = body
        ? '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + body
        : "";
      pptxCtx.addSmartArt(saKey, {
        key: saKey,
        dataModelXml,
        layout: layoutId,
        style: styleId,
        color: colorId,
      });
    }

    const x = convertToEmu(opts.x ?? 0);
    const y = convertToEmu(opts.y ?? 0);
    const w = convertToEmu(opts.width ?? "100px");
    const h = convertToEmu(opts.height ?? "100px");

    const parts: string[] = [];

    // p:nvGraphicFramePr
    parts.push(
      `<p:nvGraphicFramePr><p:cNvPr id="${id}" name="${escapeXml(name)}"/>` +
        `<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>` +
        `<p:nvPr/></p:nvGraphicFramePr>`,
    );

    // p:xfrm
    parts.push(`<p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></p:xfrm>`);

    // a:graphic > a:graphicData > dgm:relIds (placeholders)
    parts.push(
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">` +
        `<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
        `r:dm="{smartart:${saKey}}" r:lo="{smartart-lo:${saKey}}" r:qs="{smartart-qs:${saKey}}" r:cs="{smartart-cs:${saKey}}"/>` +
        `</a:graphicData></a:graphic>`,
    );

    return `<p:graphicFrame>${parts.join("")}</p:graphicFrame>`;
  },

  parse(el, _ctx) {
    const result: Partial<SmartArtDescriptorOptions> = {};

    // Position from p:xfrm
    const xfrm = findChild(el, "p:xfrm");
    if (xfrm) Object.assign(result, readPositionFromXfrm(xfrm));

    // Name from p:nvGraphicFramePr → p:cNvPr
    const nvGfxFramePr = findChild(el, "p:nvGraphicFramePr");
    if (nvGfxFramePr) {
      const cNvPr = findChild(nvGfxFramePr, "p:cNvPr");
      if (cNvPr) {
        const name = attr(cNvPr, "name");
        if (name) result.name = name;
      }
    }

    // SmartArt data via dgm:relIds → r:dm
    const relIds = findDeep(el, "dgm:relIds")[0];
    if (relIds) {
      const rId = attr(relIds, "r:dm");
      if (rId) {
        const dataPath = _ctx.resolveRelationship(rId);
        if (dataPath) {
          const dataEl = _ctx.getPart(dataPath);
          if (dataEl) {
            parseSmartArtDataXml(dataEl, result);
          }
        }
      }
    }

    return result as SmartArtDescriptorOptions;
  },
};

/** Parse SmartArt data XML into options. */
function parseSmartArtDataXml(dataEl: Element, result: Partial<SmartArtDescriptorOptions>): void {
  const model = findChild(dataEl, "dgm:dataModel");
  if (!model) return;

  const pts = findChild(model, "dgm:pts");
  if (!pts) return;

  // Build node text map and extract layout/style/color from doc point
  const nodeMap = new Map<string, string>();

  for (const pt of pts.elements ?? []) {
    if (pt.name !== "dgm:pt") continue;
    const ptType = attr(pt, "type");
    const modelId = attr(pt, "modelId");

    // Document root — extract layout/style/color
    if (ptType === "doc") {
      const prSet = findChild(pt, "dgm:prSet");
      if (prSet) {
        const loTypeId = attr(prSet, "loTypeId") ?? "";
        const qsTypeId = attr(prSet, "qsTypeId") ?? "";
        const csTypeId = attr(prSet, "csTypeId") ?? "";
        const layout = loTypeId.split("/").pop();
        if (layout) result.layout = layout;
        const style = qsTypeId.split("/").pop();
        if (style) result.style = style;
        const color = csTypeId.split("/").pop();
        if (color) result.color = color;
      }
      continue;
    }

    // Skip connection points
    if (ptType === "conn") continue;

    // Node — extract text
    if (ptType === "node" && modelId) {
      const t = findDeep(pt, "a:t")[0];
      const text = t ? extractText(t) : "";
      nodeMap.set(modelId, text);
    }
  }

  // Build tree from connections
  const cxnLst = findChild(model, "dgm:cxnLst");
  if (!cxnLst) {
    result.nodes = [];
    return;
  }

  const childrenMap = new Map<string, string[]>();
  for (const cxn of cxnLst.elements ?? []) {
    if (cxn.name !== "dgm:cxn") continue;
    const srcId = attr(cxn, "srcId");
    const destId = attr(cxn, "destId");
    if (!srcId || !destId || !nodeMap.has(destId)) continue;

    let arr = childrenMap.get(srcId);
    if (!arr) {
      arr = [];
      childrenMap.set(srcId, arr);
    }
    arr.push(destId);
  }

  const topIds = childrenMap.get("0") ?? [];
  result.nodes = topIds.map((id) => buildSmartArtNode(id, nodeMap, childrenMap));
}

function extractText(t: Element): string {
  return (t.elements ?? [])
    .filter((e) => e.type === "text")
    .map((e) => String(e.text ?? ""))
    .join("");
}

function buildSmartArtNode(
  id: string,
  nodeMap: Map<string, string>,
  childrenMap: Map<string, string[]>,
): TreeNode {
  const text = nodeMap.get(id) ?? "";
  const childIds = childrenMap.get(id) ?? [];
  if (childIds.length === 0) return { text };
  return { text, children: childIds.map((cid) => buildSmartArtNode(cid, nodeMap, childrenMap)) };
}
