/**
 * SmartArt (p:graphicFrame with diagram) descriptor for PPTX.
 *
 * Produces a graphicFrame with SmartArt relationship placeholders.
 * The actual diagram data is registered in PptxWriteContext for
 * separate compilation by the compiler.
 *
 * @module
 */

import {
  convertToEmu,
  createTransform2D,
  imageTypeFromPath,
  toUint8Array,
  transform2DDesc,
} from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import type { ReadContext } from "@office-open/core/descriptor";
import { partPathToRelsPath, resolveRelationshipTarget } from "@office-open/core/opc";
import {
  type SmartArtRawParts,
  createDataModel,
  definitionId,
  isBuiltinColorDefinition,
  isBuiltinLayoutDefinition,
  isBuiltinStyleDefinition,
  parseColorDefinition,
  parseDataModelOptions,
  parseLayoutDefinition,
  parseStyleDefinition,
  stringifyDataModelOptions,
  type SmartArtRawMediaPart,
  type TreeNode,
} from "@office-open/core/smartart";
import { attr, findChild, findFirst } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { PptxWriteContext } from "../../context";
import type { SmartArtDataRelationshipOptions, SmartArtOptions } from "../smartart";
import {
  readGraphicFrameLocking,
  readGraphicFrameHyperlink,
  stringifyGraphicFrameCnvPr,
  readNvPrPlaceholder,
  stringifyCnvGraphicFramePr,
  stringifyNvPr,
} from "./graphic-frame";
import { readCnvPr } from "./shape";
import { nextSlideDrawingId } from "./slide-drawing-ids";

// ── Types ──

// ── ID counter ──

let _nextSmartArtId = 1024;

/** Register companion images without mutating the caller's SmartArt options. */
function withRegisteredRawMedia(
  pptxCtx: PptxWriteContext,
  raw: SmartArtRawParts,
): SmartArtRawParts {
  const registered = [...(raw.media ?? []), ...(raw.layoutMedia ?? [])];
  if (registered.length === 0) return raw;
  const renames = new Map<string, string>();
  for (const media of registered) {
    const data = toUint8Array(media.data);
    const placeholder = pptxCtx.addMedia(data, imageTypeFromPath(media.fileName), media.fileName);
    const outputName = placeholder.slice(1, -1);
    if (outputName !== media.fileName) renames.set(media.fileName, outputName);
  }
  if (renames.size === 0) return raw;
  const rename = (parts: SmartArtRawMediaPart[] | undefined) =>
    parts?.map((media) => {
      const outputName = renames.get(media.fileName);
      return outputName ? { ...media, fileName: outputName } : media;
    });
  return {
    ...raw,
    media: rename(raw.media),
    layoutMedia: rename(raw.layoutMedia),
  };
}

// ── SmartArt descriptor ──

export const smartArtDesc: CustomDescriptor<SmartArtOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const pptxCtx = ctx as PptxWriteContext;
    const id = opts.id ?? nextSlideDrawingId(ctx) ?? _nextSmartArtId++;
    const name = opts.name ?? `Diagram ${id}`;
    pptxCtx.registerShapeId(name, id);
    const saKey = opts.smartArtKey ?? pptxCtx.nextSmartArtKey();

    // Custom definitions embed their own id in the doc point's type ids.
    const layoutId =
      typeof opts.layout === "object" ? definitionId(opts.layout) : (opts.layout ?? "default");
    const styleId =
      typeof opts.style === "object" ? definitionId(opts.style) : (opts.style ?? "simple1");
    const colorId =
      typeof opts.color === "object" ? definitionId(opts.color) : (opts.color ?? "accent1_2");

    // Register SmartArt data with context — an explicitly empty node tree
    // still round-trips (the data part keeps its bare doc point); only an
    // absent tree skips registration.
    if (opts.nodes !== undefined || opts.dataModel) {
      const body = opts.dataModel
        ? stringifyDataModelOptions(opts.dataModel, ctx)
        : createDataModel(opts.nodes, layoutId, styleId, colorId, ctx.reproducible);
      const dataModelXml =
        body !== undefined && body.length > 0
          ? '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + body
          : "";
      pptxCtx.addSmartArt(saKey, {
        key: saKey,
        dataModelXml,
        layout: opts.layout ?? "default",
        style: opts.style ?? "simple1",
        color: opts.color ?? "accent1_2",
        ...(opts.raw ? { raw: withRegisteredRawMedia(pptxCtx, opts.raw) } : {}),
        ...(opts.sourceDataRelationships
          ? { sourceDataRelationships: opts.sourceDataRelationships }
          : {}),
        ...(opts.sourceRelIds ? { sourceRelIds: opts.sourceRelIds } : {}),
        ...(opts.sourcePaths ? { sourcePaths: opts.sourcePaths } : {}),
      });
    }

    const x = convertToEmu(opts.x ?? 0);
    const y = convertToEmu(opts.y ?? 0);
    const w = convertToEmu(opts.width ?? "100px");
    const h = convertToEmu(opts.height ?? "100px");

    const parts: string[] = [];

    // p:nvGraphicFramePr
    parts.push(
      `<p:nvGraphicFramePr>${stringifyGraphicFrameCnvPr(id, opts, name, ctx)}` +
        `${stringifyCnvGraphicFramePr(opts.locking)}` +
        `${stringifyNvPr(opts)}</p:nvGraphicFramePr>`,
    );

    // p:xfrm
    parts.push(
      createTransform2D(
        {
          x,
          y,
          width: w,
          height: h,
          ...(opts.flipHorizontal !== undefined ? { flipHorizontal: opts.flipHorizontal } : {}),
          ...(opts.flipVertical !== undefined ? { flipVertical: opts.flipVertical } : {}),
          ...(opts.rotation !== undefined ? { rotation: opts.rotation } : {}),
        },
        "p:xfrm",
      ),
    );

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
    if (xfrm) Object.assign(result, transform2DDesc.parse(xfrm, _ctx));

    // Name from p:nvGraphicFramePr → p:cNvPr
    Object.assign(result, readCnvPr(el, "p:nvGraphicFramePr"));
    const locking = readGraphicFrameLocking(findChild(el, "p:nvGraphicFramePr"), _ctx);
    const hyperlink = readGraphicFrameHyperlink(findChild(el, "p:nvGraphicFramePr"), _ctx);
    readNvPrPlaceholder(findChild(el, "p:nvGraphicFramePr") ?? el, result);
    if (hyperlink) result.hyperlink = hyperlink;
    if (locking !== undefined) result.locking = locking;

    // SmartArt data via dgm:relIds → r:dm, plus the layout/style/colors parts
    const relIds = findFirst(el, "dgm:relIds");
    if (relIds) {
      let dataPath: string | undefined;
      const rId = attr(relIds, "r:dm");
      if (rId) {
        dataPath = _ctx.resolveRelationship(rId);
        if (dataPath) {
          const dataEl = _ctx.getPart(dataPath);
          if (dataEl) {
            const dataModel = parseDataModelOptions(
              dataEl,
              readPartRelationshipContext(dataPath, _ctx),
            );
            if (dataModel) result.dataModel = dataModel;
            parseSmartArtDataXml(dataEl, result);
          }
          const slideRels = (_ctx as { slideRelationships?: ReadonlyMap<string, string> })
            .slideRelationships;
          const drawingPaths = [...(slideRels?.values() ?? [])].filter((path) =>
            path.startsWith("ppt/diagrams/drawing"),
          );
          const rawData = readRawParts(dataPath, _ctx, drawingPaths);
          result.raw = rawData.raw;
          if (rawData.relationships.length > 0)
            result.sourceDataRelationships = rawData.relationships;
        }
      }

      const layoutPath = readRelatedPartPath(_ctx, relIds, "r:lo");
      const layoutEl = layoutPath ? _ctx.getPart(layoutPath) : undefined;
      if (layoutPath && layoutEl) {
        const layout = parseLayoutDefinition(
          layoutEl,
          readPartRelationshipContext(layoutPath, _ctx),
        );
        const id = layout.uniqueId?.split("/").pop();
        result.layout = id && isBuiltinLayoutDefinition(id, layout) ? id : layout;
        const layoutMedia = readDiagramMedia(layoutPath, _ctx);
        if (layoutMedia.length > 0) result.raw = { ...result.raw, layoutMedia };
      }
      const stylePath = readRelatedPartPath(_ctx, relIds, "r:qs");
      const styleEl = stylePath ? _ctx.getPart(stylePath) : undefined;
      if (styleEl) {
        const style = parseStyleDefinition(styleEl);
        const id = style.uniqueId?.split("/").pop();
        result.style = id && isBuiltinStyleDefinition(id, style) ? id : style;
      }
      const colorPath = readRelatedPartPath(_ctx, relIds, "r:cs");
      const colorEl = colorPath ? _ctx.getPart(colorPath) : undefined;
      if (colorEl) {
        const color = parseColorDefinition(colorEl);
        const id = color.uniqueId?.split("/").pop();
        result.color = id && isBuiltinColorDefinition(id, color) ? id : color;
      }

      result.sourceRelIds = {
        dm: attr(relIds, "r:dm") ?? undefined,
        lo: attr(relIds, "r:lo") ?? undefined,
        qs: attr(relIds, "r:qs") ?? undefined,
        cs: attr(relIds, "r:cs") ?? undefined,
      };
      result.sourcePaths = {
        data: dataPath ?? undefined,
        layout: layoutPath ?? undefined,
        quickStyle: stylePath ?? undefined,
        colors: colorPath ?? undefined,
      };
    }

    return result as SmartArtOptions;
  },
};

/** Resolve a dgm:relIds relationship attribute to its part path. */
function readRelatedPartPath(
  ctx: {
    resolveRelationship(rId: string): string | undefined;
  },
  relIds: Element,
  attrName: "r:lo" | "r:qs" | "r:cs",
) {
  const rId = attr(relIds, attrName);
  if (!rId) return undefined;
  return ctx.resolveRelationship(rId);
}

function readPartRelationshipContext(partPath: string, parent: ReadContext): ReadContext {
  const relsEl = parent.getPart(partPathToRelsPath(partPath));
  const relationships = new Map<string, string>();
  for (const relationship of relsEl?.elements ?? []) {
    if (relationship.name !== "Relationship") continue;
    const id = attr(relationship, "Id");
    const target = attr(relationship, "Target");
    if (!id || !target || attr(relationship, "TargetMode") === "External") continue;
    relationships.set(id, resolveRelationshipTarget(partPath, target));
  }
  return {
    resolveRelationship: (rId) => relationships.get(rId),
    getPart: (path) => parent.getPart(path),
    getRaw: (path) => parent.getRaw(path),
  };
}

/** Collect verbatim diagram parts behind one graphicFrame. */
function readRawParts(
  dataPath: string,
  ctx: {
    getPart(path: string): Element | undefined;
    getRaw(path: string): Uint8Array | undefined;
  },
  drawingPaths: readonly string[],
): { raw?: SmartArtRawParts; relationships: SmartArtDataRelationshipOptions[] } {
  const raw: SmartArtRawParts = {};
  let drawingRelationship = false;
  const relsPath = partPathToRelsPath(dataPath);
  const relsEl = ctx.getPart(relsPath);
  if (relsEl) {
    const media = new Map<string, { fileName: string; data: Uint8Array }>();
    for (const rel of relsEl.elements ?? []) {
      if (rel.name !== "Relationship") continue;
      const type = attr(rel, "Type") ?? "";
      const target = attr(rel, "Target");
      if (!target || attr(rel, "TargetMode") === "External") continue;
      const path = resolveRelationshipTarget(dataPath, target);
      if (type.endsWith("/diagramDrawing")) {
        drawingRelationship = true;
        raw.drawing = ctx.getRaw(path);
      }
      if (!type.endsWith("/image")) continue;
      const bytes = ctx.getRaw(path);
      const fileName = path.split("/").pop() ?? path;
      if (bytes && !media.has(fileName)) media.set(fileName, { fileName, data: bytes });
    }
    if (media.size > 0) raw.media = [...media.values()];
  }
  const directDrawingPath = drawingPaths.find((path) => ctx.getRaw(path) !== undefined);
  const dataRelsDrawingPath = (relsEl?.elements ?? [])
    .filter(
      (rel) => rel.name === "Relationship" && (attr(rel, "Type") ?? "").endsWith("/diagramDrawing"),
    )
    .map((rel) => resolveRelationshipTarget(dataPath, attr(rel, "Target") ?? ""))
    .find((path) => ctx.getRaw(path) !== undefined);
  const drawingPath = directDrawingPath ?? dataRelsDrawingPath;
  if (drawingPath) {
    raw.drawing = ctx.getRaw(drawingPath);
    raw.drawingFileName = drawingPath.split("/").pop();
  } else {
    const index = dataPath.match(/\/data(\d+)\.xml$/)?.[1];
    if (index !== undefined) raw.drawing = ctx.getRaw(`ppt/diagrams/drawing${index}.xml`);
  }
  raw.drawingRelationship = drawingRelationship;
  return { raw, relationships: readDataRelationships(relsEl) };
}

/** Preserve every source data-part relationship, including hyperlink rels. */
function readDataRelationships(relsEl: Element | undefined) {
  const relationships: {
    id: string;
    relationshipType: string;
    target: string;
    targetMode?: "External";
  }[] = [];
  for (const rel of relsEl?.elements ?? []) {
    if (rel.name !== "Relationship") continue;
    const id = attr(rel, "Id");
    const relationshipType = attr(rel, "Type");
    const target = attr(rel, "Target");
    if (!id || !relationshipType || !target) continue;
    relationships.push({
      id,
      relationshipType,
      target,
      ...(attr(rel, "TargetMode") === "External" ? { targetMode: "External" as const } : {}),
    });
  }
  return relationships;
}

/** Images referenced by one diagram part's own relationships. */
function readDiagramMedia(partPath: string, parent: ReadContext): SmartArtRawMediaPart[] {
  const relsEl = parent.getPart(partPathToRelsPath(partPath));
  const media = new Map<string, SmartArtRawMediaPart>();
  for (const relationship of relsEl?.elements ?? []) {
    if (relationship.name !== "Relationship") continue;
    const target = attr(relationship, "Target");
    if (!target || attr(relationship, "TargetMode") === "External") continue;
    if (!(attr(relationship, "Type") ?? "").endsWith("/image")) continue;
    const path = resolveRelationshipTarget(partPath, target);
    const data = parent.getRaw(path);
    const fileName = path.split("/").pop() ?? path;
    if (data && !media.has(path)) media.set(path, { fileName, data, partPath: path });
  }
  return [...media.values()];
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
  let docModelId: string | undefined;

  for (const pt of pts.elements ?? []) {
    if (pt.name !== "dgm:pt") continue;
    const ptType = attr(pt, "type"); // XSD default: "node"
    const modelId = attr(pt, "modelId");

    // Document root — extract layout/style/color
    if (ptType === "doc") {
      docModelId = modelId;
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

    // Nodes — the XSD default type. pres/asst/parTrans/sibTrans points carry
    // no user text.
    if ((ptType === undefined || ptType === "node") && modelId) {
      const t = findFirst(pt, "a:t");
      const text = t ? extractText(t) : "";
      nodeMap.set(modelId, text);
    }
  }

  // Build tree from data connections (XSD default cxn type: "parOf";
  // presOf/presParOf are presentation-tree edges that would pollute it)
  const cxnLst = findChild(model, "dgm:cxnLst");
  if (!cxnLst) {
    result.nodes = [];
    return;
  }

  const childrenMap = new Map<string, string[]>();
  for (const cxn of cxnLst.elements ?? []) {
    if (cxn.name !== "dgm:cxn") continue;
    const cxnType = attr(cxn, "type");
    if (cxnType !== undefined && cxnType !== "parOf") continue;
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

  // The doc point is the tree root — referenced either by its own modelId
  // (2007-era files) or by the reserved id "0".
  const topIds = childrenMap.get(docModelId ?? "") ?? childrenMap.get("0") ?? [];
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
