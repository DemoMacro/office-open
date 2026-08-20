/**
 * XLSX Drawing — stringify helpers for spreadsheetDrawing anchors.
 *
 * @module
 */

import { convertToEmu } from "@office-open/core";
import type { UniversalMeasure } from "@office-open/core";
import type { WriteContext } from "@office-open/core/descriptor";
import {
  buildHyperlinkElement,
  connectorLockingDesc,
  createSourceRectangle,
  graphicFrameLockingDesc,
  registerHyperlink,
  pictureLockingDesc,
  groupShapePropertiesDesc,
  shapePropertiesDesc,
  stringifyBlipEffects,
  stringifyEndpointConnection,
  stringifyNonVisualDrawingProperties,
  textBodyDesc,
} from "@office-open/core/drawing";
import type {
  ConnectorLockingOptions,
  EndpointConnectionOptions,
  GraphicFrameLockingOptions,
  NonVisualDrawingPropertiesOptions,
  ShapePropertiesOptions,
  TextBodyOptions,
  TextHyperlinkOptions,
} from "@office-open/core/drawing";
import type { DefaultShapeStyleOptions } from "@office-open/core/theme";
import { stringifyShapeStyle } from "@office-open/core/theme";
import { escapeXml } from "@office-open/xml";

import type {
  DrawingAnchorOptions,
  DrawingContentPartOptions,
  ConnectorOptions,
  GroupOptions,
  DrawingPictureOptions,
  DrawingChartOptions,
  ShapeOptions,
} from "./types";
import { ANCHOR_TYPES } from "./types";

// ── Constants ──

export const XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
export const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
export const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const C_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";

export const DEFAULT_EXTENT_CX = 400000;
export const DEFAULT_EXTENT_CY = 300000;

// ── Stringify helpers ──

/** Marker cell (0-based col/row + EMU offsets). */
function markerXml(
  col: number,
  colOff: number | UniversalMeasure,
  row: number,
  rowOff: number | UniversalMeasure,
): string {
  return `<xdr:col>${col - 1}</xdr:col><xdr:colOff>${convertToEmu(colOff)}</xdr:colOff><xdr:row>${row - 1}</xdr:row><xdr:rowOff>${convertToEmu(rowOff)}</xdr:rowOff>`;
}

export function clientDataXml(obj: {
  locksWithSheet?: boolean;
  printsWithSheet?: boolean;
}): string {
  // Both attributes default to 1 — emit them only when set explicitly, so a
  // source that wrote a bare <xdr:clientData/> round-trips as-is.
  const attrs: string[] = [];
  if (obj.locksWithSheet !== undefined)
    attrs.push(`fLocksWithSheet="${obj.locksWithSheet ? 1 : 0}"`);
  if (obj.printsWithSheet !== undefined)
    attrs.push(`fPrintsWithSheet="${obj.printsWithSheet ? 1 : 0}"`);
  return attrs.length ? `<xdr:clientData ${attrs.join(" ")}/>` : "<xdr:clientData/>";
}

/** Wrap an anchored object in the appropriate xdr:*Anchor element. */
export function wrapAnchor(opts: DrawingAnchorOptions, inner: string): string {
  const anchorType = opts.anchorType ?? ANCHOR_TYPES.twoCell;
  const cx = convertToEmu(opts.extentCx ?? DEFAULT_EXTENT_CX);
  const cy = convertToEmu(opts.extentCy ?? DEFAULT_EXTENT_CY);

  if (anchorType === ANCHOR_TYPES.absolute) {
    const x = convertToEmu(opts.absoluteX ?? 0);
    const y = convertToEmu(opts.absoluteY ?? 0);
    return `<xdr:absoluteAnchor><xdr:pos x="${x}" y="${y}"/><xdr:ext cx="${cx}" cy="${cy}"/>${inner}</xdr:absoluteAnchor>`;
  }

  const from = markerXml(opts.col, opts.colOffset ?? 0, opts.row, opts.rowOffset ?? 0);

  if (anchorType === ANCHOR_TYPES.oneCell) {
    return `<xdr:oneCellAnchor><xdr:from>${from}</xdr:from><xdr:ext cx="${cx}" cy="${cy}"/>${inner}</xdr:oneCellAnchor>`;
  }

  // twoCell — editAs defaults to "twoCell" in the schema; emit the attribute
  // only when set explicitly so a source that omitted it round-trips as-is.
  const editAsAttr = opts.editAs === undefined ? "" : ` editAs="${opts.editAs}"`;
  const to = markerXml(
    opts.toCol ?? opts.col + 1,
    opts.toColOffset ?? 0,
    opts.toRow ?? opts.row + 1,
    opts.toRowOffset ?? 0,
  );
  return `<xdr:twoCellAnchor${editAsAttr}><xdr:from>${from}</xdr:from><xdr:to>${to}</xdr:to>${inner}</xdr:twoCellAnchor>`;
}

function picXml(
  img: DrawingPictureOptions,
  id: number,
  cx: number,
  cy: number,
  ctx: WriteContext,
): string {
  // Round-tripped spPr (rotation/flip/bwMode/fill) wins; fresh pictures get
  // the position-only standard form.
  const spPr =
    (img.spPr
      ? shapePropertiesDesc.stringify(img.spPr, ctx)
      : shapePropertiesDesc.stringify(
          { x: 0, y: 0, width: cx, height: cy, geometry: "rect" },
          ctx,
        )) ?? "";
  // preferRelativeResize defaults to true — emit the attribute only when the
  // source carried it explicitly.
  const prAttr =
    img.preferRelativeResize === undefined
      ? ""
      : ` preferRelativeResize="${img.preferRelativeResize ? 1 : 0}"`;
  const locks = img.locking ? (pictureLockingDesc.stringify(img.locking, ctx) ?? "") : "";
  const cNvPicPr = locks
    ? `<xdr:cNvPicPr${prAttr}>${locks}</xdr:cNvPicPr>`
    : `<xdr:cNvPicPr${prAttr}/>`;
  const effects = img.blipEffects ? stringifyBlipEffects(img.blipEffects, ctx) : "";
  // r:embed carries the embedded copy, r:link the external source — a
  // linked-only picture has no rId, a purely embedded one no linkRId.
  const blipAttrs: string[] = [];
  if (img.rId) blipAttrs.push(`r:embed="${img.rId}"`);
  if (img.linkRId) blipAttrs.push(`r:link="${img.linkRId}"`);
  const attrs = blipAttrs.join(" ");
  const open = blipAttrs.length ? `<a:blip ${attrs}` : "<a:blip";
  const blip = effects ? `${open}>${effects}</a:blip>` : `${open}/>`;
  const srcRect = img.sourceRectangle ? createSourceRectangle(img.sourceRectangle) : "";
  const bwModeAttr = img.blackWhiteMode ? ` bwMode="${img.blackWhiteMode}"` : "";
  return (
    `<xdr:pic><xdr:nvPicPr>${stringifyNonVisualDrawingProperties("xdr:cNvPr", id, img, `Picture ${id}`, hlinkClickXml(img.hyperlink, ctx))}${cNvPicPr}</xdr:nvPicPr>` +
    `<xdr:blipFill>${blip}${srcRect}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
    `<xdr:spPr${bwModeAttr}>${spPr}</xdr:spPr></xdr:pic>`
  );
}

export function stringifyImage(img: DrawingPictureOptions, id: number, ctx: WriteContext): string {
  const cx = convertToEmu(img.extentCx ?? DEFAULT_EXTENT_CX);
  const cy = convertToEmu(img.extentCy ?? DEFAULT_EXTENT_CY);
  const pic = picXml(img, id, cx, cy, ctx);
  return wrapAnchor(img, `${pic}${clientDataXml(img)}`);
}

/**
 * Build an xdr:graphicFrame hosting a chart reference. The xfrm extent is the
 * caller's choice: twoCell anchors carry 0×0 (position comes from the cell
 * markers), absolute anchors carry the real frame size.
 */
/** a:hlinkClick inside cNvPr — registers the target and returns the element. */
function hlinkClickXml(
  hyperlink: TextHyperlinkOptions | undefined,
  ctx: WriteContext | undefined,
): string | undefined {
  if (!hyperlink || !ctx) return undefined;
  return buildHyperlinkElement("a:hlinkClick", hyperlink, registerHyperlink(hyperlink, ctx));
}

export function graphicFrameXml(
  id: number,
  cNvPr: NonVisualDrawingPropertiesOptions | undefined,
  name: string,
  rId: string,
  cx: number,
  cy: number,
  ctx?: WriteContext,
  extras: {
    frameLocks?: GraphicFrameLockingOptions;
    macro?: string;
    hyperlink?: TextHyperlinkOptions;
  } = {},
): string {
  // Locks are optional in CT_NonVisualGraphicFrameProperties — emit them only
  // when the source carried them (a bare <xdr:cNvGraphicFramePr/> round-trips).
  // @macro round-trips even when empty (Word writes macro="").
  // The locking descriptor consumes no context, so stringify works even when
  // the caller (chart path) passes no ctx.
  const locks = extras.frameLocks
    ? (graphicFrameLockingDesc.stringify(extras.frameLocks, ctx as WriteContext) ?? "")
    : "";
  const cNvGraphicFramePr = locks
    ? `<xdr:cNvGraphicFramePr>${locks}</xdr:cNvGraphicFramePr>`
    : "<xdr:cNvGraphicFramePr/>";
  const macroAttr = extras.macro === undefined ? "" : ` macro="${escapeXml(extras.macro)}"`;
  return (
    `<xdr:graphicFrame${macroAttr}><xdr:nvGraphicFramePr>${stringifyNonVisualDrawingProperties("xdr:cNvPr", id, cNvPr, name, hlinkClickXml(extras.hyperlink, ctx))}` +
    `${cNvGraphicFramePr}</xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></xdr:xfrm>` +
    `<a:graphic><a:graphicData uri="${C_URI}">` +
    `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${R_NS}" r:id="${rId}"/>` +
    `</a:graphicData></a:graphic></xdr:graphicFrame>`
  );
}

export function stringifyChart(chart: DrawingChartOptions, id: number, ctx?: WriteContext): string {
  // Charts keep their historical default footprint (10 columns × 17 rows)
  // when no to corner is set; graphicFrame xfrm stays 0×0 for twoCellAnchor
  // because the position comes from the cell markers.
  const anchor = { toCol: chart.col + 9, toRow: chart.row + 16, ...chart };
  const clientData = clientDataXml(chart);
  const isTwoCell = (anchor.anchorType ?? ANCHOR_TYPES.twoCell) === ANCHOR_TYPES.twoCell;
  const cx = isTwoCell ? 0 : convertToEmu(anchor.extentCx ?? DEFAULT_EXTENT_CX);
  const cy = isTwoCell ? 0 : convertToEmu(anchor.extentCy ?? DEFAULT_EXTENT_CY);
  const frame = graphicFrameXml(id, chart, `Chart ${id}`, chart.rId, cx, cy, ctx, {
    frameLocks: chart.frameLocks,
    macro: chart.macro,
    hyperlink: chart.hyperlink,
  });
  return wrapAnchor(anchor, `${frame}${clientData}`);
}

/** Build the inner xdr:sp content (nvSpPr + spPr + optional style/txBody). */
function buildShapeContent(
  shape: NonVisualDrawingPropertiesOptions & {
    textBox?: boolean;
    hyperlink?: TextHyperlinkOptions;
  },
  id: number,
  fallbackName: string,
  spPr: ShapePropertiesOptions,
  textBody: TextBodyOptions | undefined,
  ctx: WriteContext,
  attrs = "",
  style?: DefaultShapeStyleOptions,
): string {
  const cNvPr = shape;
  const spPrXml = shapePropertiesDesc.stringify(spPr, ctx) ?? "";
  const styleXml = style ? stringifyShapeStyle(style, ctx, "xdr:style") : "";
  const txBodyXml = textBody
    ? `<xdr:txBody>${textBodyDesc.stringify(textBody, ctx)}</xdr:txBody>`
    : "";
  return `<xdr:sp${attrs}><xdr:nvSpPr>${stringifyNonVisualDrawingProperties("xdr:cNvPr", id, cNvPr, fallbackName, hlinkClickXml(shape.hyperlink, ctx))}<xdr:cNvSpPr${shape.textBox === undefined ? "" : ` txBox="${shape.textBox ? 1 : 0}"`}/></xdr:nvSpPr><xdr:spPr>${spPrXml}</xdr:spPr>${styleXml}${txBodyXml}</xdr:sp>`;
}

/** Build the inner xdr:cxnSp content (nvCxnSpPr + spPr). */
function buildConnectorContent(
  cNvPr: (NonVisualDrawingPropertiesOptions & { hyperlink?: TextHyperlinkOptions }) | undefined,
  id: number,
  fallbackName: string,
  spPr: ShapePropertiesOptions,
  ctx: WriteContext,
  attrs = "",
  connector?: {
    locking?: ConnectorLockingOptions;
    startConnection?: EndpointConnectionOptions;
    endConnection?: EndpointConnectionOptions;
    style?: DefaultShapeStyleOptions;
  },
): string {
  const spPrXml = shapePropertiesDesc.stringify(spPr, ctx) ?? "";
  const styleXml = connector?.style ? stringifyShapeStyle(connector.style, ctx, "xdr:style") : "";
  const cNvCxnSpPrInner: string[] = [];
  if (connector?.locking) {
    const locks = connectorLockingDesc.stringify(connector.locking, ctx);
    if (locks) cNvCxnSpPrInner.push(locks);
  }
  if (connector?.startConnection) {
    cNvCxnSpPrInner.push(stringifyEndpointConnection("stCxn", connector.startConnection));
  }
  if (connector?.endConnection) {
    cNvCxnSpPrInner.push(stringifyEndpointConnection("endCxn", connector.endConnection));
  }
  const cNvCxnSpPr = cNvCxnSpPrInner.length
    ? `<xdr:cNvCxnSpPr>${cNvCxnSpPrInner.join("")}</xdr:cNvCxnSpPr>`
    : "<xdr:cNvCxnSpPr/>";
  return `<xdr:cxnSp${attrs}><xdr:nvCxnSpPr>${stringifyNonVisualDrawingProperties("xdr:cNvPr", id, cNvPr, fallbackName, hlinkClickXml(cNvPr?.hyperlink, ctx))}${cNvCxnSpPr}</xdr:nvCxnSpPr><xdr:spPr>${spPrXml}</xdr:spPr>${styleXml}</xdr:cxnSp>`;
}

export function stringifyShape(shape: ShapeOptions, id: number, ctx: WriteContext): string {
  const xml = buildShapeContent(
    shape,
    id,
    `Shape ${id}`,
    shape.spPr,
    shape.textBody,
    ctx,
    macroTextlinkAttrs(shape),
    shape.style,
  );
  return wrapAnchor(shape, `${xml}${clientDataXml(shape)}`);
}

export function stringifyConnector(conn: ConnectorOptions, id: number, ctx: WriteContext): string {
  const xml = buildConnectorContent(
    conn,
    id,
    `Connector ${id}`,
    conn.spPr,
    ctx,
    macroTextlinkAttrs(conn),
    conn,
  );
  return wrapAnchor(conn, `${xml}${clientDataXml(conn)}`);
}

export function stringifyContentPart(cp: DrawingContentPartOptions): string {
  return wrapAnchor(cp, `<xdr:contentPart r:id="${cp.rId}"/>${clientDataXml(cp)}`);
}

/** Build xdr:grpSp content and return the next available cNvPr id. */
export function buildGroup(
  grp: GroupOptions,
  id: number,
  ctx: WriteContext,
): { xml: string; nextId: number } {
  const grpSpPrXml = groupShapePropertiesDesc.stringify(grp.grpSpPr, ctx) ?? "";
  let childId = id + 1;
  const children: string[] = [];
  for (const childShape of grp.shapes ?? []) {
    children.push(
      buildShapeContent(
        childShape,
        childId,
        `Shape ${childId}`,
        childShape.spPr,
        childShape.textBody,
        ctx,
        macroTextlinkAttrs(childShape),
        childShape.style,
      ),
    );
    childId++;
  }
  for (const childConn of grp.connectors ?? []) {
    children.push(
      buildConnectorContent(
        childConn,
        childId,
        `Connector ${childId}`,
        childConn.spPr,
        ctx,
        macroTextlinkAttrs(childConn),
        childConn,
      ),
    );
    childId++;
  }
  const xml =
    `<xdr:grpSp><xdr:nvGrpSpPr>${stringifyNonVisualDrawingProperties("xdr:cNvPr", id, grp, `Group ${id}`, hlinkClickXml(grp.hyperlink, ctx))}<xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>` +
    `<xdr:grpSpPr>${grpSpPrXml}</xdr:grpSpPr>${children.join("")}</xdr:grpSp>`;
  return { xml, nextId: childId };
}

/** CT_Shape attribute string (macro/textlink) with leading space, or empty. */
function macroTextlinkAttrs(shape: { macro?: string; textlink?: string }): string {
  const a: string[] = [];
  if (shape.macro !== undefined) a.push(`macro="${escapeXml(shape.macro)}"`);
  if (shape.textlink !== undefined) a.push(`textlink="${escapeXml(shape.textlink)}"`);
  return a.length ? " " + a.join(" ") : "";
}
