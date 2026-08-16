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
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { stringifyNonVisualDrawingProperties } from "@office-open/core/drawing";
import {
  COLOR_CATEGORIES,
  LAYOUT_CATEGORIES,
  STYLE_CATEGORIES,
  createDataModel,
  definitionId,
  parseColorDefinition,
  parseLayoutDefinition,
  parseStyleDefinition,
  type TreeNode,
} from "@office-open/core/smartart";
import { attr, findChild, findFirst } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { PptxWriteContext } from "../../context";
import type { SmartArtOptions } from "../smartart";
import { readCnvPr, readPositionFromXfrm } from "./shape";

// ── Types ──

// ── ID counter ──

let _nextSmartArtId = 1024;

// ── SmartArt descriptor ──

export const smartArtDesc: CustomDescriptor<SmartArtOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const pptxCtx = ctx as PptxWriteContext;
    const id = opts.id ?? _nextSmartArtId++;
    const name = opts.name ?? `Diagram ${id}`;
    const saKey = opts.smartArtKey ?? pptxCtx.nextSmartArtKey();

    // Custom definitions embed their own id in the doc point's type ids.
    const layoutId =
      typeof opts.layout === "object" ? definitionId(opts.layout) : (opts.layout ?? "default");
    const styleId =
      typeof opts.style === "object" ? definitionId(opts.style) : (opts.style ?? "simple1");
    const colorId =
      typeof opts.color === "object" ? definitionId(opts.color) : (opts.color ?? "accent1_2");

    // Register SmartArt data with context
    if (opts.nodes && opts.nodes.length > 0) {
      const body = createDataModel(opts.nodes, layoutId, styleId, colorId);
      const dataModelXml = body
        ? '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + body
        : "";
      pptxCtx.addSmartArt(saKey, {
        key: saKey,
        dataModelXml,
        layout: opts.layout ?? "default",
        style: opts.style ?? "simple1",
        color: opts.color ?? "accent1_2",
      });
    }

    const x = convertToEmu(opts.x ?? 0);
    const y = convertToEmu(opts.y ?? 0);
    const w = convertToEmu(opts.width ?? "100px");
    const h = convertToEmu(opts.height ?? "100px");

    const parts: string[] = [];

    // p:nvGraphicFramePr
    parts.push(
      `<p:nvGraphicFramePr>${stringifyNonVisualDrawingProperties("p:cNvPr", id, opts, name)}` +
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
    const result: Partial<SmartArtOptions> = {};

    // Position from p:xfrm
    const xfrm = findChild(el, "p:xfrm");
    if (xfrm) Object.assign(result, readPositionFromXfrm(xfrm));

    // Name from p:nvGraphicFramePr → p:cNvPr
    Object.assign(result, readCnvPr(el, "p:nvGraphicFramePr"));

    // SmartArt data via dgm:relIds → r:dm, plus the layout/style/colors parts
    const relIds = findFirst(el, "dgm:relIds");
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

      // Custom definitions come back structured; built-in stubs fold to their
      // id string so round-tripping a built-in diagram keeps the compact form.
      const layoutEl = readRelatedPart(_ctx, relIds, "r:lo");
      if (layoutEl) {
        const layout = parseLayoutDefinition(layoutEl);
        const id = layout.uniqueId?.split("/").pop();
        result.layout = id && id in LAYOUT_CATEGORIES ? id : layout;
      }
      const styleEl = readRelatedPart(_ctx, relIds, "r:qs");
      if (styleEl) {
        const style = parseStyleDefinition(styleEl);
        const id = style.uniqueId?.split("/").pop();
        result.style = id && id in STYLE_CATEGORIES ? id : style;
      }
      const colorEl = readRelatedPart(_ctx, relIds, "r:cs");
      if (colorEl) {
        const color = parseColorDefinition(colorEl);
        const id = color.uniqueId?.split("/").pop();
        result.color = id && id in COLOR_CATEGORIES ? id : color;
      }
    }

    return result as SmartArtOptions;
  },
};

/** Resolve a dgm:relIds relationship attribute to its parsed part element. */
function readRelatedPart(
  ctx: {
    resolveRelationship(rId: string): string | undefined;
    getPart(path: string): Element | undefined;
  },
  relIds: Element,
  attrName: "r:lo" | "r:qs" | "r:cs",
): Element | undefined {
  const rId = attr(relIds, attrName);
  if (!rId) return undefined;
  const path = ctx.resolveRelationship(rId);
  return path ? ctx.getPart(path) : undefined;
}

/** Parse SmartArt data XML into options. */
function parseSmartArtDataXml(dataEl: Element, result: Partial<SmartArtOptions>): void {
  // The data part's root element IS dgm:dataModel (CT_DataModel: ptLst + cxnLst).
  const model = dataEl.name === "dgm:dataModel" ? dataEl : findChild(dataEl, "dgm:dataModel");
  if (!model) return;

  const pts = findChild(model, "dgm:ptLst");
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
      const t = findFirst(pt, "a:t");
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
