/**
 * XLSX Drawing — anchor objects and descriptor.
 *
 * Generates xl/drawings/drawing{n}.xml using the spreadsheetDrawing
 * namespace for anchoring drawing objects (images, charts, shapes, groups,
 * connectors, content parts) to worksheet cells.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import type { Element as XmlElement } from "@office-open/xml";

import {
  parseChartAnchor,
  parseConnectorAnchor,
  parseContentPartAnchor,
  parseGroupAnchor,
  parseImageAnchor,
  parseShapeAnchor,
  findXdr,
} from "./parse";
import {
  buildGroup,
  clientDataXml,
  stringifyChart,
  stringifyConnector,
  stringifyContentPart,
  stringifyImage,
  stringifyShape,
  wrapAnchor,
  A_NS,
  R_NS,
  XDR_NS,
} from "./stringify";
import type {
  DrawingChartOptions,
  DrawingContentPartOptions,
  ConnectorOptions,
  GroupOptions,
  DrawingPictureOptions,
  DrawingOptions,
  ShapeOptions,
} from "./types";

// ── Descriptor ──

/** Original cNvPr id of an anchored object (the anchor's first cNvPr is the object's own — its nv*Pr container precedes everything else). */
function readAnchorShapeId(anchor: XmlElement): number | undefined {
  const firstCnvPr = (el: XmlElement): XmlElement | undefined => {
    for (const child of el.elements ?? []) {
      if ((child.name ?? "").endsWith("cNvPr")) return child;
      const found = firstCnvPr(child);
      if (found) return found;
    }
    return undefined;
  };
  const cNvPr = firstCnvPr(anchor);
  if (!cNvPr) return undefined;
  const id = Number(cNvPr.attributes?.["id"]);
  return Number.isNaN(id) ? undefined : id;
}

export const drawingDesc: CustomDescriptor<DrawingOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const images = opts.images ?? [];
    const charts = opts.charts ?? [];
    const shapes = opts.shapes ?? [];
    const connectors = opts.connectors ?? [];
    const groups = opts.groups ?? [];
    const contentParts = opts.contentParts ?? [];
    const total =
      images.length +
      charts.length +
      shapes.length +
      connectors.length +
      groups.length +
      contentParts.length;
    if (total === 0) return undefined;

    // Anchors carry their document order from parse (the z-order when objects
    // overlap). Emit in that order; fresh options without one keep the
    // per-type bucket order (stable sort).
    interface Emission {
      order: number | undefined;
      shapeId: number | undefined;
      emit: (id: number) => string;
    }
    const emissions: Emission[] = [];
    for (const img of images)
      emissions.push({
        order: img.zOrder,
        shapeId: img.shapeId,
        emit: (id) => stringifyImage(img, id, ctx),
      });
    for (const chart of charts)
      emissions.push({
        order: chart.zOrder,
        shapeId: chart.shapeId,
        emit: (id) => stringifyChart(chart, id, ctx),
      });
    for (const shape of shapes)
      emissions.push({
        order: shape.zOrder,
        shapeId: shape.shapeId,
        emit: (id) => stringifyShape(shape, id, ctx),
      });
    for (const conn of connectors)
      emissions.push({
        order: conn.zOrder,
        shapeId: conn.shapeId,
        emit: (id) => stringifyConnector(conn, id, ctx),
      });
    for (const grp of groups)
      emissions.push({
        order: grp.zOrder,
        shapeId: grp.shapeId,
        emit: (id) => wrapAnchor(grp, `${buildGroup(grp, id, ctx).xml}${clientDataXml(grp)}`),
      });
    for (const cp of contentParts)
      emissions.push({
        order: cp.zOrder,
        shapeId: cp.shapeId,
        emit: () => stringifyContentPart(cp),
      });
    emissions.sort((a, b) => (a.order ?? Infinity) - (b.order ?? Infinity));

    // Ids: keep each object's original cNvPr id when present (round-trip);
    // allocate unused ones for the rest (fresh authoring).
    const usedIds = new Set<number>(
      emissions.map((e) => e.shapeId).filter((n): n is number => n !== undefined),
    );
    const p: string[] = [`<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">`];
    let next = 1;
    for (const e of emissions) {
      const id =
        e.shapeId ??
        ((): number => {
          while (usedIds.has(next)) next++;
          usedIds.add(next);
          return next;
        })();
      p.push(e.emit(id));
    }

    p.push("</xdr:wsDr>");
    return p.join("");
  },

  parse(el, ctx) {
    const result: Partial<DrawingOptions> = {};
    const images: DrawingPictureOptions[] = [];
    const charts: DrawingChartOptions[] = [];
    const shapes: ShapeOptions[] = [];
    const connectors: ConnectorOptions[] = [];
    const groups: GroupOptions[] = [];
    const contentParts: DrawingContentPartOptions[] = [];

    // Document order of the anchors is the z-order when objects overlap —
    // record it so stringify can re-emit them interleaved instead of grouped
    // per type bucket.
    let order = 0;
    const stamp = <T extends { zOrder?: number; shapeId?: number }>(
      obj: T,
      anchor: XmlElement,
    ): T => {
      obj.zOrder = order++;
      const shapeId = readAnchorShapeId(anchor);
      if (shapeId !== undefined) obj.shapeId = shapeId;
      return obj;
    };

    for (const anchor of el.elements ?? []) {
      // Office writes spreadsheetDrawing anchors in the default namespace or
      // with the xdr: prefix — normalize before dispatching.
      const rawName = anchor.name ?? "";
      const name = rawName.startsWith("xdr:") ? rawName.slice(4) : rawName;
      if (name !== "twoCellAnchor" && name !== "oneCellAnchor" && name !== "absoluteAnchor") {
        continue;
      }

      const pic = findXdr(anchor, "pic");
      if (pic) {
        images.push(stamp(parseImageAnchor(anchor, pic, name, ctx), anchor));
        continue;
      }

      const graphicFrame = findXdr(anchor, "graphicFrame");
      if (graphicFrame) {
        const chart = parseChartAnchor(anchor, graphicFrame, name, ctx);
        if (chart) charts.push(stamp(chart, anchor));
        continue;
      }

      const sp = findXdr(anchor, "sp");
      if (sp) {
        shapes.push(stamp(parseShapeAnchor(anchor, sp, name, ctx), anchor));
        continue;
      }

      const cxnSp = findXdr(anchor, "cxnSp");
      if (cxnSp) {
        connectors.push(stamp(parseConnectorAnchor(anchor, cxnSp, name, ctx), anchor));
        continue;
      }

      const grpSp = findXdr(anchor, "grpSp");
      if (grpSp) {
        groups.push(stamp(parseGroupAnchor(anchor, grpSp, name, ctx), anchor));
        continue;
      }

      const contentPart = findXdr(anchor, "contentPart");
      if (contentPart) {
        const cp = parseContentPartAnchor(anchor, contentPart, name);
        if (cp) contentParts.push(stamp(cp, anchor));
      }
    }

    if (images.length > 0) result.images = images;
    if (charts.length > 0) result.charts = charts;
    if (shapes.length > 0) result.shapes = shapes;
    if (connectors.length > 0) result.connectors = connectors;
    if (groups.length > 0) result.groups = groups;
    if (contentParts.length > 0) result.contentParts = contentParts;
    return result as DrawingOptions;
  },
};
