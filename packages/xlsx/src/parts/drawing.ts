/**
 * XLSX Drawing — image and chart anchor types and descriptor.
 *
 * Generates xl/drawings/drawing{n}.xml using the spreadsheetDrawing
 * namespace for anchoring images and charts to worksheet cells.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

// ── Types (used by compiler) ──

export interface ImageOptions {
  /** 1-based column */
  col: number;
  /** Column offset in EMU (default 0) */
  colOffset?: number;
  /** 1-based row */
  row: number;
  /** Row offset in EMU (default 0) */
  rowOffset?: number;
  /** Relationship ID for the image */
  rId: string;
  /** Lock anchor with sheet (default true) */
  locksWithSheet?: boolean;
  /** Print with sheet (default true) */
  printsWithSheet?: boolean;
}

export interface ChartAnchorOptions {
  /** 1-based column */
  col: number;
  /** Column offset in EMU (default 0) */
  colOffset?: number;
  /** 1-based row */
  row: number;
  /** Row offset in EMU (default 0) */
  rowOffset?: number;
  /** Relationship ID for the chart */
  rId: string;
  /** Lock anchor with sheet (default true) */
  locksWithSheet?: boolean;
  /** Print with sheet (default true) */
  printsWithSheet?: boolean;
}

// ── Descriptor Types ──

export interface DrawingImageOptions {
  /** 1-based column */
  col: number;
  /** Column offset in EMU (default 0) */
  colOffset?: number;
  /** 1-based row */
  row: number;
  /** Row offset in EMU (default 0) */
  rowOffset?: number;
  /** Relationship ID for the image */
  rId: string;
  /** Lock anchor with sheet (default true) */
  locksWithSheet?: boolean;
  /** Print with sheet (default true) */
  printsWithSheet?: boolean;
}

export interface DrawingChartOptions {
  /** 1-based column */
  col: number;
  /** Column offset in EMU (default 0) */
  colOffset?: number;
  /** 1-based row */
  row: number;
  /** Row offset in EMU (default 0) */
  rowOffset?: number;
  /** Relationship ID for the chart */
  rId: string;
  /** Lock anchor with sheet (default true) */
  locksWithSheet?: boolean;
  /** Print with sheet (default true) */
  printsWithSheet?: boolean;
}

export interface DrawingOptions {
  images?: DrawingImageOptions[];
  charts?: DrawingChartOptions[];
}

// ── Constants ──

const XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const C_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";

// ── Descriptor ──

export const drawingDesc: CustomDescriptor<DrawingOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const images = opts.images ?? [];
    const charts = opts.charts ?? [];
    if (images.length === 0 && charts.length === 0) return undefined;

    const p: string[] = [`<wsDr xmlns="${XDR_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">`];
    let id = 1;

    for (const img of images) {
      p.push(
        `<twoCellAnchor editAs="oneCell"><from><col>${img.col - 1}</col><colOff>${img.colOffset ?? 0}</colOff><row>${img.row - 1}</row><rowOff>${img.rowOffset ?? 0}</rowOff></from>`,
        `<to><col>${img.col}</col><colOff>0</colOff><row>${img.row}</row><rowOff>0</rowOff></to>`,
        `<pic><nvPicPr><cNvPr id="${id}" name="Picture ${id}"/><cNvPicPr preferRelativeResize="1"/></nvPicPr>`,
        `<blipFill><a:blip r:embed="${img.rId}"/><a:stretch><a:fillRect/></a:stretch></blipFill>`,
        `<spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="400000" cy="300000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></spPr></pic>`,
        `<clientData fLocksWithSheet="${img.locksWithSheet !== false ? 1 : 0}" fPrintsWithSheet="${img.printsWithSheet !== false ? 1 : 0}"/></twoCellAnchor>`,
      );
      id++;
    }

    for (const chart of charts) {
      p.push(
        `<twoCellAnchor editAs="oneCell"><from><col>${chart.col - 1}</col><colOff>${chart.colOffset ?? 0}</colOff><row>${chart.row - 1}</row><rowOff>${chart.rowOffset ?? 0}</rowOff></from>`,
        `<to><col>${chart.col + 8}</col><colOff>0</colOff><row>${chart.row + 15}</row><rowOff>0</rowOff></to>`,
        `<graphicFrame><nvGraphicFramePr><cNvPr id="${id}" name="Chart ${id}"/><cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></cNvGraphicFramePr></nvGraphicFramePr>`,
        `<xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xfrm>`,
        `<a:graphic><a:graphicData uri="${C_URI}"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${R_NS}" r:id="${chart.rId}"/></a:graphicData></a:graphic></graphicFrame>`,
        `<clientData fLocksWithSheet="${chart.locksWithSheet !== false ? 1 : 0}" fPrintsWithSheet="${chart.printsWithSheet !== false ? 1 : 0}"/></twoCellAnchor>`,
      );
      id++;
    }

    p.push("</wsDr>");
    return p.join("");
  },

  parse(el, _ctx) {
    const result: Record<string, unknown> = {};
    const images: DrawingImageOptions[] = [];
    const charts: DrawingChartOptions[] = [];

    for (const anchor of el.elements ?? []) {
      if (anchor.name !== "twoCellAnchor") continue;
      const from = findChild(anchor, "from");
      findChild(anchor, "to"); // consumed but not used
      if (!from) continue;

      const col = readNumChild(from, "col") + 1;
      const colOffset = readNumChild(from, "colOff") || undefined;
      const row = readNumChild(from, "row") + 1;
      const rowOffset = readNumChild(from, "rowOff") || undefined;

      // Check if this is a picture or graphicFrame
      const pic = findChild(anchor, "pic");
      if (pic) {
        const blip = findChild(findChild(pic, "blipFill") ?? pic, "a:blip");
        const rId = blip?.attributes?.["r:embed"] as string | undefined;
        if (rId) {
          const clientData = findChild(anchor, "clientData");
          images.push({
            col,
            colOffset,
            row,
            rowOffset,
            rId,
            locksWithSheet: clientData?.attributes?.["fLocksWithSheet"] !== "0",
            printsWithSheet: clientData?.attributes?.["fPrintsWithSheet"] !== "0",
          });
        }
        continue;
      }

      const graphicFrame = findChild(anchor, "graphicFrame");
      if (graphicFrame) {
        const graphicData = findChild(
          findChild(graphicFrame, "a:graphic") ?? graphicFrame,
          "a:graphicData",
        );
        const chartEl = graphicData ? findChild(graphicData, "c:chart") : undefined;
        const rId = chartEl?.attributes?.["r:id"] as string | undefined;
        if (rId) {
          const clientData = findChild(anchor, "clientData");
          charts.push({
            col,
            colOffset,
            row,
            rowOffset,
            rId,
            locksWithSheet: clientData?.attributes?.["fLocksWithSheet"] !== "0",
            printsWithSheet: clientData?.attributes?.["fPrintsWithSheet"] !== "0",
          });
        }
      }
    }

    if (images.length > 0) result.images = images;
    if (charts.length > 0) result.charts = charts;
    return result as Record<string, unknown>;
  },
};

// ── Helpers ──

function readNumChild(el: XmlElement, tag: string): number {
  const child = findChild(el, tag);
  if (!child?.elements?.length) return 0;
  const n = Number(child.elements[0]?.text ?? "");
  return Number.isNaN(n) ? 0 : n;
}
