/**
 * XLSX Drawing — image and chart anchor types and descriptor.
 *
 * Generates xl/drawings/drawing{n}.xml using the spreadsheetDrawing
 * namespace for anchoring images and charts to worksheet cells.
 *
 * @module
 */

import { convertToEmu } from "@office-open/core";
import type { UniversalMeasure } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

// ── Types (used by compiler) ──

export interface ImageOptions {
  /** 1-based column */
  col: number;
  /** Column offset in EMU (default 0) */
  colOffset?: number | UniversalMeasure;
  /** 1-based row */
  row: number;
  /** Row offset in EMU (default 0) */
  rowOffset?: number | UniversalMeasure;
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
  colOffset?: number | UniversalMeasure;
  /** 1-based row */
  row: number;
  /** Row offset in EMU (default 0) */
  rowOffset?: number | UniversalMeasure;
  /** Relationship ID for the chart */
  rId: string;
  /** Lock anchor with sheet (default true) */
  locksWithSheet?: boolean;
  /** Print with sheet (default true) */
  printsWithSheet?: boolean;
}

// ── Descriptor Types ──

/** How a drawing is anchored to the worksheet (xdr:*Anchor element). */
export const ANCHOR_TYPES = {
  twoCell: "twoCell",
  oneCell: "oneCell",
  absolute: "absolute",
} as const;
export type AnchorType = (typeof ANCHOR_TYPES)[keyof typeof ANCHOR_TYPES];

/** editAs behavior for twoCellAnchor (ST_EditAs). */
export const EDIT_AS_TYPES = {
  twoCell: "twoCell",
  oneCell: "oneCell",
  absolute: "absolute",
} as const;
export type EditAsType = (typeof EDIT_AS_TYPES)[keyof typeof EDIT_AS_TYPES];

export interface DrawingImageOptions {
  /** 1-based column (from marker) */
  col: number;
  /** Column offset in EMU (default 0) */
  colOffset?: number | UniversalMeasure;
  /** 1-based row (from marker) */
  row: number;
  /** Row offset in EMU (default 0) */
  rowOffset?: number | UniversalMeasure;
  /** To cell column (1-based) for twoCellAnchor. Defaults to col + 1. */
  toCol?: number;
  /** To cell row (1-based) for twoCellAnchor. Defaults to row + 1. */
  toRow?: number;
  /** To cell column offset in EMU. */
  toColOffset?: number | UniversalMeasure;
  /** To cell row offset in EMU. */
  toRowOffset?: number | UniversalMeasure;
  /** Relationship ID for the image */
  rId: string;
  /** Anchor type (default "twoCell"). */
  anchorType?: AnchorType;
  /** editAs for twoCellAnchor (default "oneCell"). */
  editAs?: EditAsType;
  /** Absolute X in EMU (absoluteAnchor). */
  absoluteX?: number | UniversalMeasure;
  /** Absolute Y in EMU (absoluteAnchor). */
  absoluteY?: number | UniversalMeasure;
  /** Image width in EMU (a:ext cx, default 400000). */
  extentCx?: number | UniversalMeasure;
  /** Image height in EMU (a:ext cy, default 300000). */
  extentCy?: number | UniversalMeasure;
  /** Lock anchor with sheet (default true) */
  locksWithSheet?: boolean;
  /** Print with sheet (default true) */
  printsWithSheet?: boolean;
}

export interface DrawingChartOptions {
  /** 1-based column */
  col: number;
  /** Column offset in EMU (default 0) */
  colOffset?: number | UniversalMeasure;
  /** 1-based row */
  row: number;
  /** Row offset in EMU (default 0) */
  rowOffset?: number | UniversalMeasure;
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

const DEFAULT_EXTENT_CX = 400000;
const DEFAULT_EXTENT_CY = 300000;

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
      p.push(stringifyImage(img, id));
      id++;
    }

    for (const chart of charts) {
      p.push(stringifyChart(chart, id));
      id++;
    }

    p.push("</wsDr>");
    return p.join("");
  },

  parse(el, _ctx) {
    const result: Partial<DrawingOptions> = {};
    const images: DrawingImageOptions[] = [];
    const charts: DrawingChartOptions[] = [];

    for (const anchor of el.elements ?? []) {
      const name = anchor.name;
      if (name !== "twoCellAnchor" && name !== "oneCellAnchor" && name !== "absoluteAnchor") {
        continue;
      }

      const pic = findChild(anchor, "pic");
      if (pic) {
        images.push(parseImageAnchor(anchor, pic, name));
        continue;
      }

      const graphicFrame = findChild(anchor, "graphicFrame");
      if (graphicFrame) {
        const chart = parseChartAnchor(anchor, graphicFrame);
        if (chart) charts.push(chart);
      }
    }

    if (images.length > 0) result.images = images;
    if (charts.length > 0) result.charts = charts;
    return result as DrawingOptions;
  },
};

// ── Stringify helpers ──

/** Marker cell (0-based col/row + EMU offsets). */
function markerXml(
  col: number,
  colOff: number | UniversalMeasure,
  row: number,
  rowOff: number | UniversalMeasure,
): string {
  return `<col>${col - 1}</col><colOff>${convertToEmu(colOff)}</colOff><row>${row - 1}</row><rowOff>${convertToEmu(rowOff)}</rowOff>`;
}

function clientDataXml(img: { locksWithSheet?: boolean; printsWithSheet?: boolean }): string {
  const locks = img.locksWithSheet !== false ? 1 : 0;
  const prints = img.printsWithSheet !== false ? 1 : 0;
  return `<clientData fLocksWithSheet="${locks}" fPrintsWithSheet="${prints}"/>`;
}

function picXml(rId: string, id: number, cx: number, cy: number): string {
  return (
    `<pic><nvPicPr><cNvPr id="${id}" name="Picture ${id}"/><cNvPicPr preferRelativeResize="1"/></nvPicPr>` +
    `<blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></blipFill>` +
    `<spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></spPr></pic>`
  );
}

function stringifyImage(img: DrawingImageOptions, id: number): string {
  const anchorType = img.anchorType ?? ANCHOR_TYPES.twoCell;
  const cx = convertToEmu(img.extentCx ?? DEFAULT_EXTENT_CX);
  const cy = convertToEmu(img.extentCy ?? DEFAULT_EXTENT_CY);
  const pic = picXml(img.rId, id, cx, cy);
  const clientData = clientDataXml(img);

  if (anchorType === ANCHOR_TYPES.absolute) {
    const x = convertToEmu(img.absoluteX ?? 0);
    const y = convertToEmu(img.absoluteY ?? 0);
    return (
      `<absoluteAnchor><pos x="${x}" y="${y}"/><ext cx="${cx}" cy="${cy}"/>` +
      `${pic}${clientData}</absoluteAnchor>`
    );
  }

  const from = markerXml(img.col, img.colOffset ?? 0, img.row, img.rowOffset ?? 0);

  if (anchorType === ANCHOR_TYPES.oneCell) {
    return (
      `<oneCellAnchor><from>${from}</from><ext cx="${cx}" cy="${cy}"/>` +
      `${pic}${clientData}</oneCellAnchor>`
    );
  }

  // twoCell
  const editAs = img.editAs ?? EDIT_AS_TYPES.oneCell;
  const to = markerXml(
    img.toCol ?? img.col + 1,
    img.toColOffset ?? 0,
    img.toRow ?? img.row + 1,
    img.toRowOffset ?? 0,
  );
  return (
    `<twoCellAnchor editAs="${editAs}"><from>${from}</from><to>${to}</to>` +
    `${pic}${clientData}</twoCellAnchor>`
  );
}

function stringifyChart(chart: DrawingChartOptions, id: number): string {
  // Charts default to twoCellAnchor (existing behavior).
  const from = markerXml(chart.col, chart.colOffset ?? 0, chart.row, chart.rowOffset ?? 0);
  const to = markerXml(chart.col + 9, 0, chart.row + 16, 0);
  const clientData = clientDataXml(chart);
  return (
    `<twoCellAnchor editAs="oneCell"><from>${from}</from><to>${to}</to>` +
    `<graphicFrame><nvGraphicFramePr><cNvPr id="${id}" name="Chart ${id}"/>` +
    `<cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></cNvGraphicFramePr></nvGraphicFramePr>` +
    `<xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xfrm>` +
    `<a:graphic><a:graphicData uri="${C_URI}">` +
    `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${R_NS}" r:id="${chart.rId}"/>` +
    `</a:graphicData></a:graphic></graphicFrame>${clientData}</twoCellAnchor>`
  );
}

// ── Parse helpers ──

function readNumChild(el: XmlElement, tag: string): number {
  const child = findChild(el, tag);
  if (!child?.elements?.length) return 0;
  const n = Number(child.elements[0]?.text ?? "");
  return Number.isNaN(n) ? 0 : n;
}

interface Marker {
  col: number;
  colOffset?: number | UniversalMeasure;
  row: number;
  rowOffset?: number | UniversalMeasure;
}

function readMarker(el: XmlElement): Marker {
  return {
    col: readNumChild(el, "col") + 1,
    colOffset: readNumChild(el, "colOff") || undefined,
    row: readNumChild(el, "row") + 1,
    rowOffset: readNumChild(el, "rowOff") || undefined,
  };
}

function readPicRId(pic: XmlElement): string | undefined {
  const blipFill = findChild(pic, "blipFill") ?? pic;
  const blip = findChild(blipFill, "a:blip");
  return blip?.attributes?.["r:embed"] as string | undefined;
}

/** Picture extent from pic/spPr/a:xfrm/a:ext (actual image size in EMU). */
function readPicExtent(pic: XmlElement): { cx?: number; cy?: number } {
  const spPr = findChild(pic, "spPr");
  const xfrm = spPr ? findChild(spPr, "a:xfrm") : undefined;
  const ext = xfrm ? findChild(xfrm, "a:ext") : undefined;
  if (!ext?.attributes) return {};
  const cx = Number(ext.attributes["cx"]);
  const cy = Number(ext.attributes["cy"]);
  return {
    cx: Number.isNaN(cx) ? undefined : cx,
    cy: Number.isNaN(cy) ? undefined : cy,
  };
}

function parseImageAnchor(anchor: XmlElement, pic: XmlElement, name: string): DrawingImageOptions {
  const result: DrawingImageOptions = {
    col: 1,
    row: 1,
    rId: readPicRId(pic) ?? "",
  };

  // Actual image size (applies to all anchor types).
  const ext = readPicExtent(pic);
  if (ext.cx !== undefined) result.extentCx = ext.cx;
  if (ext.cy !== undefined) result.extentCy = ext.cy;

  // Client data (lock/print flags) — applies to all anchor types.
  const clientData = findChild(anchor, "clientData");
  if (clientData?.attributes) {
    if (clientData.attributes["fLocksWithSheet"] !== undefined) {
      result.locksWithSheet = clientData.attributes["fLocksWithSheet"] !== "0";
    }
    if (clientData.attributes["fPrintsWithSheet"] !== undefined) {
      result.printsWithSheet = clientData.attributes["fPrintsWithSheet"] !== "0";
    }
  }

  if (name === "absoluteAnchor") {
    result.anchorType = ANCHOR_TYPES.absolute;
    const pos = findChild(anchor, "pos");
    if (pos?.attributes) {
      const x = Number(pos.attributes["x"]);
      const y = Number(pos.attributes["y"]);
      if (!Number.isNaN(x)) result.absoluteX = x;
      if (!Number.isNaN(y)) result.absoluteY = y;
    }
    return result;
  }

  const from = findChild(anchor, "from");
  if (from) {
    const m = readMarker(from);
    result.col = m.col;
    result.row = m.row;
    if (m.colOffset !== undefined) result.colOffset = m.colOffset;
    if (m.rowOffset !== undefined) result.rowOffset = m.rowOffset;
  }

  if (name === "oneCellAnchor") {
    result.anchorType = ANCHOR_TYPES.oneCell;
    return result;
  }

  // twoCellAnchor
  result.anchorType = ANCHOR_TYPES.twoCell;
  const to = findChild(anchor, "to");
  if (to) {
    const m = readMarker(to);
    result.toCol = m.col;
    result.toRow = m.row;
    if (m.colOffset !== undefined) result.toColOffset = m.colOffset;
    if (m.rowOffset !== undefined) result.toRowOffset = m.rowOffset;
  }
  const editAs = anchor.attributes?.["editAs"] as EditAsType | undefined;
  if (editAs) result.editAs = editAs;

  return result;
}

function parseChartAnchor(
  anchor: XmlElement,
  graphicFrame: XmlElement,
): DrawingChartOptions | undefined {
  const graphicData = findChild(
    findChild(graphicFrame, "a:graphic") ?? graphicFrame,
    "a:graphicData",
  );
  const chartEl = graphicData ? findChild(graphicData, "c:chart") : undefined;
  const rId = chartEl?.attributes?.["r:id"] as string | undefined;
  if (!rId) return undefined;

  const result: DrawingChartOptions = { col: 1, row: 1, rId };
  const from = findChild(anchor, "from");
  if (from) {
    const m = readMarker(from);
    result.col = m.col;
    result.row = m.row;
    if (m.colOffset !== undefined) result.colOffset = m.colOffset;
    if (m.rowOffset !== undefined) result.rowOffset = m.rowOffset;
  }
  const clientData = findChild(anchor, "clientData");
  if (clientData?.attributes) {
    if (clientData.attributes["fLocksWithSheet"] !== undefined) {
      result.locksWithSheet = clientData.attributes["fLocksWithSheet"] !== "0";
    }
    if (clientData.attributes["fPrintsWithSheet"] !== undefined) {
      result.printsWithSheet = clientData.attributes["fPrintsWithSheet"] !== "0";
    }
  }
  return result;
}
