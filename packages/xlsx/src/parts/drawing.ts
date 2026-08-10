/**
 * XLSX Drawing — anchor objects and descriptor.
 *
 * Generates xl/drawings/drawing{n}.xml using the spreadsheetDrawing
 * namespace for anchoring drawing objects (images, charts, shapes, groups,
 * connectors, content parts) to worksheet cells.
 *
 * @module
 */

import { convertToEmu } from "@office-open/core";
import type { UniversalMeasure } from "@office-open/core";
import type { CustomDescriptor, ReadContext, WriteContext } from "@office-open/core/descriptor";
import { shapePropertiesDesc, textBodyDesc } from "@office-open/core/drawingml";
import type {
  GroupTransform2DOptions,
  ShapePropertiesOptions,
  TextBodyOptions,
} from "@office-open/core/drawingml";
import { escapeXml, findChild } from "@office-open/xml";
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

/** Shared anchor fields for all anchored drawing objects. */
export interface DrawingAnchorOptions {
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
  /** Anchor type (default "twoCell"). */
  anchorType?: AnchorType;
  /** editAs for twoCellAnchor (default "oneCell"). */
  editAs?: EditAsType;
  /** Absolute X in EMU (absoluteAnchor). */
  absoluteX?: number | UniversalMeasure;
  /** Absolute Y in EMU (absoluteAnchor). */
  absoluteY?: number | UniversalMeasure;
  /** Anchor extent width in EMU (oneCell/absoluteAnchor ext, default 400000). */
  extentCx?: number | UniversalMeasure;
  /** Anchor extent height in EMU (oneCell/absoluteAnchor ext, default 300000). */
  extentCy?: number | UniversalMeasure;
  /** Lock anchor with sheet (default true) */
  locksWithSheet?: boolean;
  /** Print with sheet (default true) */
  printsWithSheet?: boolean;
}

export interface DrawingImageOptions extends DrawingAnchorOptions {
  /** Relationship ID for the image */
  rId: string;
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

/** Anchored shape (xdr:sp): geometry + optional text body. */
export interface DrawingShapeOptions extends DrawingAnchorOptions {
  /** Shape name (cNvPr name). Defaults to "Shape <id>". */
  name?: string;
  /** Shape properties (a:CT_ShapeProperties). */
  spPr: ShapePropertiesOptions;
  /** Text body (a:CT_TextBody). */
  textBody?: TextBodyOptions;
  /** macro attribute (CT_Shape). */
  macro?: string;
  /** textlink attribute (CT_Shape). */
  textlink?: string;
}

/** Anchored connector (xdr:cxnSp): line/arrow geometry via spPr. */
export interface DrawingConnectorOptions extends DrawingAnchorOptions {
  /** Connector name. Defaults to "Connector <id>". */
  name?: string;
  /** Shape properties (a:CT_ShapeProperties, typically prstGeom="line"). */
  spPr: ShapePropertiesOptions;
  /** macro attribute (CT_Connector). */
  macro?: string;
}

/** Shape nested inside a group (no anchor — positioned via spPr.xfrm). */
export interface GroupShapeChildOptions {
  name?: string;
  spPr: ShapePropertiesOptions;
  textBody?: TextBodyOptions;
  macro?: string;
  textlink?: string;
}

/** Connector nested inside a group (no anchor). */
export interface GroupConnectorChildOptions {
  name?: string;
  spPr: ShapePropertiesOptions;
  macro?: string;
}

/** Anchored group (xdr:grpSp): group transform + nested shapes/connectors. */
export interface DrawingGroupOptions extends DrawingAnchorOptions {
  /** Group name. Defaults to "Group <id>". */
  name?: string;
  /** Group shape properties (a:CT_GroupShapeProperties: group xfrm + fill/ln). */
  grpSpPr: GroupTransform2DOptions;
  /** Nested shapes. */
  shapes?: GroupShapeChildOptions[];
  /** Nested connectors. */
  connectors?: GroupConnectorChildOptions[];
}

/** Anchored external content reference (xdr:contentPart, r:id only). */
export interface DrawingContentPartOptions extends DrawingAnchorOptions {
  /** Relationship ID for the external content. */
  rId: string;
}

export interface DrawingOptions {
  images?: DrawingImageOptions[];
  charts?: DrawingChartOptions[];
  shapes?: DrawingShapeOptions[];
  connectors?: DrawingConnectorOptions[];
  groups?: DrawingGroupOptions[];
  contentParts?: DrawingContentPartOptions[];
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
    const images: DrawingImageOptions[] = [];
    const charts: DrawingChartOptions[] = [];
    const shapes: DrawingShapeOptions[] = [];
    const connectors: DrawingConnectorOptions[] = [];
    const groups: DrawingGroupOptions[] = [];
    const contentParts: DrawingContentPartOptions[] = [];

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
        continue;
      }

      const sp = findChild(anchor, "sp");
      if (sp) {
        shapes.push(parseShapeAnchor(anchor, sp, name, ctx));
        continue;
      }

      const cxnSp = findChild(anchor, "cxnSp");
      if (cxnSp) {
        connectors.push(parseConnectorAnchor(anchor, cxnSp, name, ctx));
        continue;
      }

      const grpSp = findChild(anchor, "grpSp");
      if (grpSp) {
        groups.push(parseGroupAnchor(anchor, grpSp, name, ctx));
        continue;
      }

      const contentPart = findChild(anchor, "contentPart");
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

function clientDataXml(obj: { locksWithSheet?: boolean; printsWithSheet?: boolean }): string {
  const locks = obj.locksWithSheet !== false ? 1 : 0;
  const prints = obj.printsWithSheet !== false ? 1 : 0;
  return `<clientData fLocksWithSheet="${locks}" fPrintsWithSheet="${prints}"/>`;
}

/** Wrap an anchored object in the appropriate xdr:*Anchor element. */
function wrapAnchor(opts: DrawingAnchorOptions, inner: string): string {
  const anchorType = opts.anchorType ?? ANCHOR_TYPES.twoCell;
  const cx = convertToEmu(opts.extentCx ?? DEFAULT_EXTENT_CX);
  const cy = convertToEmu(opts.extentCy ?? DEFAULT_EXTENT_CY);

  if (anchorType === ANCHOR_TYPES.absolute) {
    const x = convertToEmu(opts.absoluteX ?? 0);
    const y = convertToEmu(opts.absoluteY ?? 0);
    return `<absoluteAnchor><pos x="${x}" y="${y}"/><ext cx="${cx}" cy="${cy}"/>${inner}</absoluteAnchor>`;
  }

  const from = markerXml(opts.col, opts.colOffset ?? 0, opts.row, opts.rowOffset ?? 0);

  if (anchorType === ANCHOR_TYPES.oneCell) {
    return `<oneCellAnchor><from>${from}</from><ext cx="${cx}" cy="${cy}"/>${inner}</oneCellAnchor>`;
  }

  // twoCell
  const editAs = opts.editAs ?? EDIT_AS_TYPES.oneCell;
  const to = markerXml(
    opts.toCol ?? opts.col + 1,
    opts.toColOffset ?? 0,
    opts.toRow ?? opts.row + 1,
    opts.toRowOffset ?? 0,
  );
  return `<twoCellAnchor editAs="${editAs}"><from>${from}</from><to>${to}</to>${inner}</twoCellAnchor>`;
}

function picXml(rId: string, id: number, cx: number, cy: number, ctx: WriteContext): string {
  const spPr =
    shapePropertiesDesc.stringify({ x: 0, y: 0, width: cx, height: cy, geometry: "rect" }, ctx) ??
    "";
  return (
    `<pic><nvPicPr><cNvPr id="${id}" name="Picture ${id}"/><cNvPicPr preferRelativeResize="1"/></nvPicPr>` +
    `<blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></blipFill>` +
    `<spPr>${spPr}</spPr></pic>`
  );
}

function stringifyImage(img: DrawingImageOptions, id: number, ctx: WriteContext): string {
  const cx = convertToEmu(img.extentCx ?? DEFAULT_EXTENT_CX);
  const cy = convertToEmu(img.extentCy ?? DEFAULT_EXTENT_CY);
  const pic = picXml(img.rId, id, cx, cy, ctx);
  return wrapAnchor(img, `${pic}${clientDataXml(img)}`);
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

/** Build the inner xdr:sp content (nvSpPr + spPr + optional txBody). */
function buildShapeContent(
  name: string,
  id: number,
  spPr: ShapePropertiesOptions,
  textBody: TextBodyOptions | undefined,
  ctx: WriteContext,
  attrs = "",
): string {
  const spPrXml = shapePropertiesDesc.stringify(spPr, ctx) ?? "";
  const txBodyXml = textBody ? `<txBody>${textBodyDesc.stringify(textBody, ctx)}</txBody>` : "";
  return `<sp${attrs}><nvSpPr><cNvPr id="${id}" name="${escapeXml(name)}"/><cNvSpPr/></nvSpPr><spPr>${spPrXml}</spPr>${txBodyXml}</sp>`;
}

/** Build the inner xdr:cxnSp content (nvCxnSpPr + spPr). */
function buildConnectorContent(
  name: string,
  id: number,
  spPr: ShapePropertiesOptions,
  ctx: WriteContext,
  attrs = "",
): string {
  const spPrXml = shapePropertiesDesc.stringify(spPr, ctx) ?? "";
  return `<cxnSp${attrs}><nvCxnSpPr><cNvPr id="${id}" name="${escapeXml(name)}"/><cNvCxnSpPr/></nvCxnSpPr><spPr>${spPrXml}</spPr></cxnSp>`;
}

function stringifyShape(shape: DrawingShapeOptions, id: number, ctx: WriteContext): string {
  const xml = buildShapeContent(
    shape.name ?? `Shape ${id}`,
    id,
    shape.spPr,
    shape.textBody,
    ctx,
    macroTextlinkAttrs(shape),
  );
  return wrapAnchor(shape, `${xml}${clientDataXml(shape)}`);
}

function stringifyConnector(conn: DrawingConnectorOptions, id: number, ctx: WriteContext): string {
  const xml = buildConnectorContent(
    conn.name ?? `Connector ${id}`,
    id,
    conn.spPr,
    ctx,
    macroTextlinkAttrs(conn),
  );
  return wrapAnchor(conn, `${xml}${clientDataXml(conn)}`);
}

function stringifyContentPart(cp: DrawingContentPartOptions): string {
  return wrapAnchor(cp, `<contentPart r:id="${cp.rId}"/>${clientDataXml(cp)}`);
}

/** Build xdr:grpSp content and return the next available cNvPr id. */
function buildGroup(
  grp: DrawingGroupOptions,
  id: number,
  ctx: WriteContext,
): { xml: string; nextId: number } {
  const grpSpPrXml = shapePropertiesDesc.stringify(grp.grpSpPr, ctx) ?? "";
  let childId = id + 1;
  const children: string[] = [];
  for (const childShape of grp.shapes ?? []) {
    children.push(
      buildShapeContent(
        childShape.name ?? `Shape ${childId}`,
        childId,
        childShape.spPr,
        childShape.textBody,
        ctx,
        macroTextlinkAttrs(childShape),
      ),
    );
    childId++;
  }
  for (const childConn of grp.connectors ?? []) {
    children.push(
      buildConnectorContent(
        childConn.name ?? `Connector ${childId}`,
        childId,
        childConn.spPr,
        ctx,
        macroTextlinkAttrs(childConn),
      ),
    );
    childId++;
  }
  const xml =
    `<grpSp><nvGrpSpPr><cNvPr id="${id}" name="${escapeXml(grp.name ?? `Group ${id}`)}"/><cNvGrpSpPr/></nvGrpSpPr>` +
    `<grpSpPr>${grpSpPrXml}</grpSpPr>${children.join("")}</grpSp>`;
  return { xml, nextId: childId };
}

/** CT_Shape attribute string (macro/textlink) with leading space, or empty. */
function macroTextlinkAttrs(shape: { macro?: string; textlink?: string }): string {
  const a: string[] = [];
  if (shape.macro !== undefined) a.push(`macro="${escapeXml(shape.macro)}"`);
  if (shape.textlink !== undefined) a.push(`textlink="${escapeXml(shape.textlink)}"`);
  return a.length ? " " + a.join(" ") : "";
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

/** Populate anchor fields (col/row/extent/anchorType/editAs/locks) from an anchor element. */
function readAnchorFields(anchor: XmlElement, name: string, result: DrawingAnchorOptions): void {
  result.col = 1;
  result.row = 1;

  const clientData = findChild(anchor, "clientData");
  if (clientData?.attributes) {
    if (clientData.attributes["fLocksWithSheet"] !== undefined) {
      result.locksWithSheet = clientData.attributes["fLocksWithSheet"] !== "0";
    }
    if (clientData.attributes["fPrintsWithSheet"] !== undefined) {
      result.printsWithSheet = clientData.attributes["fPrintsWithSheet"] !== "0";
    }
  }

  const readExt = (el: XmlElement | undefined): void => {
    if (!el?.attributes) return;
    const cx = Number(el.attributes["cx"]);
    const cy = Number(el.attributes["cy"]);
    if (!Number.isNaN(cx)) result.extentCx = cx;
    if (!Number.isNaN(cy)) result.extentCy = cy;
  };

  if (name === "absoluteAnchor") {
    result.anchorType = ANCHOR_TYPES.absolute;
    const pos = findChild(anchor, "pos");
    if (pos?.attributes) {
      const x = Number(pos.attributes["x"]);
      const y = Number(pos.attributes["y"]);
      if (!Number.isNaN(x)) result.absoluteX = x;
      if (!Number.isNaN(y)) result.absoluteY = y;
    }
    readExt(findChild(anchor, "ext"));
    return;
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
    readExt(findChild(anchor, "ext"));
    return;
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
}

/** Read cNvPr name from a non-visual properties child (nvSpPr/nvCxnSpPr/nvGrpSpPr). */
function readCNvName(parent: XmlElement, nonVisualTag: string): string | undefined {
  const nonVisual = findChild(parent, nonVisualTag);
  const cNvPr = nonVisual ? findChild(nonVisual, "cNvPr") : undefined;
  return cNvPr?.attributes?.["name"] !== undefined ? String(cNvPr.attributes["name"]) : undefined;
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

  readAnchorFields(anchor, name, result);
  return result;
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

function parseShapeAnchor(
  anchor: XmlElement,
  sp: XmlElement,
  name: string,
  ctx: ReadContext,
): DrawingShapeOptions {
  const result = { col: 1, row: 1, spPr: {} } as DrawingShapeOptions;
  readAnchorFields(anchor, name, result);

  const cNvName = readCNvName(sp, "nvSpPr");
  if (cNvName !== undefined) result.name = cNvName;

  const spPr = findChild(sp, "spPr");
  if (spPr) result.spPr = shapePropertiesDesc.parse(spPr, ctx);

  const txBody = findChild(sp, "txBody");
  if (txBody) result.textBody = textBodyDesc.parse(txBody, ctx);

  if (sp.attributes?.["macro"] !== undefined) result.macro = String(sp.attributes["macro"]);
  if (sp.attributes?.["textlink"] !== undefined)
    result.textlink = String(sp.attributes["textlink"]);
  return result;
}

function parseConnectorAnchor(
  anchor: XmlElement,
  cxnSp: XmlElement,
  name: string,
  ctx: ReadContext,
): DrawingConnectorOptions {
  const result = { col: 1, row: 1, spPr: {} } as DrawingConnectorOptions;
  readAnchorFields(anchor, name, result);

  const cNvName = readCNvName(cxnSp, "nvCxnSpPr");
  if (cNvName !== undefined) result.name = cNvName;

  const spPr = findChild(cxnSp, "spPr");
  if (spPr) result.spPr = shapePropertiesDesc.parse(spPr, ctx);

  if (cxnSp.attributes?.["macro"] !== undefined) result.macro = String(cxnSp.attributes["macro"]);
  return result;
}

function parseGroupAnchor(
  anchor: XmlElement,
  grpSp: XmlElement,
  name: string,
  ctx: ReadContext,
): DrawingGroupOptions {
  const result = { col: 1, row: 1, grpSpPr: {} } as DrawingGroupOptions;
  readAnchorFields(anchor, name, result);

  const cNvName = readCNvName(grpSp, "nvGrpSpPr");
  if (cNvName !== undefined) result.name = cNvName;

  const grpSpPrEl = findChild(grpSp, "grpSpPr");
  if (grpSpPrEl) {
    result.grpSpPr = shapePropertiesDesc.parse(grpSpPrEl, ctx) as GroupTransform2DOptions;
  }

  const shapes: GroupShapeChildOptions[] = [];
  const childConnectors: GroupConnectorChildOptions[] = [];
  for (const child of grpSp.elements ?? []) {
    if (child.name === "sp") {
      const spPr = findChild(child, "spPr");
      const childShape = {
        spPr: spPr ? shapePropertiesDesc.parse(spPr, ctx) : {},
      } as GroupShapeChildOptions;
      const childName = readCNvName(child, "nvSpPr");
      if (childName !== undefined) childShape.name = childName;
      const txBody = findChild(child, "txBody");
      if (txBody) childShape.textBody = textBodyDesc.parse(txBody, ctx);
      if (child.attributes?.["macro"] !== undefined)
        childShape.macro = String(child.attributes["macro"]);
      if (child.attributes?.["textlink"] !== undefined)
        childShape.textlink = String(child.attributes["textlink"]);
      shapes.push(childShape);
    } else if (child.name === "cxnSp") {
      const spPr = findChild(child, "spPr");
      const childConn = {
        spPr: spPr ? shapePropertiesDesc.parse(spPr, ctx) : {},
      } as GroupConnectorChildOptions;
      const childName = readCNvName(child, "nvCxnSpPr");
      if (childName !== undefined) childConn.name = childName;
      if (child.attributes?.["macro"] !== undefined)
        childConn.macro = String(child.attributes["macro"]);
      childConnectors.push(childConn);
    }
  }
  if (shapes.length > 0) result.shapes = shapes;
  if (childConnectors.length > 0) result.connectors = childConnectors;
  return result;
}

function parseContentPartAnchor(
  anchor: XmlElement,
  contentPart: XmlElement,
  name: string,
): DrawingContentPartOptions | undefined {
  const rId = contentPart.attributes?.["r:id"] as string | undefined;
  if (!rId) return undefined;
  const result = { col: 1, row: 1, rId } as DrawingContentPartOptions;
  readAnchorFields(anchor, name, result);
  return result;
}
