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

    const p: string[] = [`<wsDr xmlns="${XDR_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">`];
    let id = 1;

    for (const img of images) {
      p.push(stringifyImage(img, id, ctx));
      id++;
    }
    for (const chart of charts) {
      p.push(stringifyChart(chart, id));
      id++;
    }
    for (const shape of shapes) {
      p.push(stringifyShape(shape, id, ctx));
      id++;
    }
    for (const conn of connectors) {
      p.push(stringifyConnector(conn, id, ctx));
      id++;
    }
    for (const grp of groups) {
      const built = buildGroup(grp, id, ctx);
      p.push(wrapAnchor(grp, `${built.xml}${clientDataXml(grp)}`));
      id = built.nextId;
    }
    for (const cp of contentParts) {
      p.push(stringifyContentPart(cp));
      id++;
    }

    p.push("</wsDr>");
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
        images.push(parseImageAnchor(anchor, pic, name));
        continue;
      }

      const graphicFrame = findXdr(anchor, "graphicFrame");
      if (graphicFrame) {
        const chart = parseChartAnchor(anchor, graphicFrame, name);
        if (chart) charts.push(chart);
        continue;
      }

      const sp = findXdr(anchor, "sp");
      if (sp) {
        shapes.push(parseShapeAnchor(anchor, sp, name, ctx));
        continue;
      }

      const cxnSp = findXdr(anchor, "cxnSp");
      if (cxnSp) {
        connectors.push(parseConnectorAnchor(anchor, cxnSp, name, ctx));
        continue;
      }

      const grpSp = findXdr(anchor, "grpSp");
      if (grpSp) {
        groups.push(parseGroupAnchor(anchor, grpSp, name, ctx));
        continue;
      }

      const contentPart = findXdr(anchor, "contentPart");
      if (contentPart) {
        const cp = parseContentPartAnchor(anchor, contentPart, name);
        if (cp) contentParts.push(cp);
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
