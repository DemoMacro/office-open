/**
 * XLSX Drawing — stringify helpers for spreadsheetDrawing anchors.
 *
 * @module
 */

import { convertToEmu, emitAngle } from "@office-open/core";
import type { UniversalMeasure } from "@office-open/core";
import type { WriteContext } from "@office-open/core/descriptor";
import {
  buildHyperlinkElement,
  connectorLockingDesc,
  createSourceRectangle,
  graphicFrameLockingDesc,
  registerHyperlink,
  pictureLockingDesc,
  groupLockingDesc,
  groupShapePropertiesDesc,
  shapeLockingDesc,
  shapePropertiesDesc,
  stringifyBlipEffects,
  stringifyEndpointConnection,
  stringifyNonVisualDrawingProperties,
  textBodyDesc,
} from "@office-open/core/drawing";
import type {
  BlackWhiteMode,
  ConnectorLockingOptions,
  EndpointConnectionOptions,
  GraphicFrameLockingOptions,
  NonVisualDrawingPropertiesOptions,
  ShapeLockingOptions,
  ShapePropertiesOptions,
  TextBodyOptions,
  TextHyperlinkOptions,
} from "@office-open/core/drawing";
import type { DefaultShapeStyleOptions } from "@office-open/core/theme";
import { stringifyShapeStyle } from "@office-open/core/theme";
import { attrsRaw, escapeXml } from "@office-open/xml";

import type {
  DrawingAnchorOptions,
  DrawingContentPartOptions,
  ConnectorOptions,
  GroupOptions,
  GroupPictureChildOptions,
  DrawingPictureOptions,
  DrawingChartOptions,
  DrawingWebExtensionOptions,
  DrawingSmartArtOptions,
  ShapeOptions,
} from "./types";
import { ANCHOR_TYPES } from "./types";

// ── Constants ──

export const XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
export const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
export const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const C_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";
/** graphicData uri for the diagram (SmartArt) payload. */
export const DGM_URI = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
export const WE_URI = "http://schemas.microsoft.com/office/webextensions/webextension/2010/11";
const WE_NS = WE_URI;
const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const A14_NS = "http://schemas.microsoft.com/office/drawing/2010/main";

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
  img: DrawingPictureOptions | GroupPictureChildOptions,
  id: number,
  cx: number,
  cy: number,
  ctx: WriteContext,
): string {
  // Round-tripped spPr (rotation/flip/bwMode/fill) wins; fresh pictures get
  // the position-only standard form.
  const spPr =
    (img.properties
      ? shapePropertiesDesc.stringify(img.properties, ctx)
      : shapePropertiesDesc.stringify(
          { x: 0, y: 0, width: cx, height: cy, geometry: "rect" },
          ctx,
        )) ?? "";
  // preferRelativeResize defaults to true — emit the attribute only when the
  // source carried it explicitly.
  const prAttr =
    img.preferRelativeResize === undefined
      ? ""
      : ` preferRelativeResize="${img.preferRelativeResize ? "true" : "false"}"`;
  const locks = img.locking ? (pictureLockingDesc.stringify(img.locking, ctx) ?? "") : "";
  const prExt = img.cNvPicPrExt ? `<a:extLst>${img.cNvPicPrExt}</a:extLst>` : "";
  const cNvPicPr =
    locks || prExt
      ? `<xdr:cNvPicPr${prAttr}>${locks}${prExt}</xdr:cNvPicPr>`
      : `<xdr:cNvPicPr${prAttr}/>`;
  const effects = img.blipEffects ? stringifyBlipEffects(img.blipEffects, ctx) : "";
  // The blip extension list: the verbatim channel subsumes useLocalDpi.
  const extLst = img.blipExt
    ? `<a:extLst>${img.blipExt}</a:extLst>`
    : img.useLocalDpi !== undefined
      ? `<a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="${img.useLocalDpi ? 1 : 0}"/></a:ext></a:extLst>`
      : "";
  // r:embed carries the embedded copy, r:link the external source — a
  // linked-only picture has no rId, a purely embedded one no linkRId.
  const blipAttrs: string[] = [];
  if (img.rId) blipAttrs.push(`r:embed="${img.rId}"`);
  if (img.compression !== undefined) blipAttrs.push(`cstate="${img.compression}"`);
  if (img.linkRId) blipAttrs.push(`r:link="${img.linkRId}"`);
  const attrs = blipAttrs.join(" ");
  const open = blipAttrs.length ? `<a:blip ${attrs}` : "<a:blip";
  const blipContent = effects + extLst;
  const blip = blipContent ? `${open}>${blipContent}</a:blip>` : `${open}/>`;
  const srcRect = img.sourceRectangle ? createSourceRectangle(img.sourceRectangle) : "";
  const blipFillAttrs =
    (img.dpi !== undefined ? ` dpi="${img.dpi}"` : "") +
    (img.rotWithShape !== undefined ? ` rotWithShape="${img.rotWithShape ? 1 : 0}"` : "");
  const bwModeAttr = img.blackWhiteMode ? ` bwMode="${img.blackWhiteMode}"` : "";
  const publishedAttr = publishedObjectAttrs(img);
  return (
    `<xdr:pic${publishedAttr}><xdr:nvPicPr>${stringifyNonVisualDrawingProperties("xdr:cNvPr", id, img, `Picture ${id}`, hlinkClickXml(img.hyperlink, ctx))}${cNvPicPr}</xdr:nvPicPr>` +
    `<xdr:blipFill${blipFillAttrs}>${blip}${srcRect}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
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
    fPublished?: boolean;
    frameRotation?: number;
    frameFlipHorizontal?: boolean;
    frameFlipVertical?: boolean;
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
  const objectAttrs = `${macroAttr(extras.macro)}${publishedObjectAttrs(extras)}`;
  return (
    `<xdr:graphicFrame${objectAttrs}><xdr:nvGraphicFramePr>${stringifyNonVisualDrawingProperties("xdr:cNvPr", id, cNvPr, name, hlinkClickXml(extras.hyperlink, ctx))}` +
    `${cNvGraphicFramePr}</xdr:nvGraphicFramePr>` +
    `<xdr:xfrm${attrsRaw({
      rot: extras.frameRotation !== undefined ? emitAngle(extras.frameRotation) : undefined,
      flipH:
        extras.frameFlipHorizontal !== undefined ? (extras.frameFlipHorizontal ? 1 : 0) : undefined,
      flipV:
        extras.frameFlipVertical !== undefined ? (extras.frameFlipVertical ? 1 : 0) : undefined,
    })}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></xdr:xfrm>` +
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
  const anchorType = anchor.anchorType ?? ANCHOR_TYPES.twoCell;
  const cellAnchored = anchorType === ANCHOR_TYPES.twoCell || anchorType === ANCHOR_TYPES.oneCell;
  const cx =
    anchor.frameExtentCx !== undefined
      ? anchor.frameExtentCx
      : cellAnchored
        ? 0
        : convertToEmu(anchor.extentCx ?? DEFAULT_EXTENT_CX);
  const cy =
    anchor.frameExtentCy !== undefined
      ? anchor.frameExtentCy
      : cellAnchored
        ? 0
        : convertToEmu(anchor.extentCy ?? DEFAULT_EXTENT_CY);
  const frame = graphicFrameXml(id, chart, `Chart ${id}`, chart.rId, cx, cy, ctx, {
    frameLocks: chart.frameLocks,
    macro: chart.macro,
    hyperlink: chart.hyperlink,
    fPublished: chart.fPublished,
    frameRotation: chart.frameRotation,
    frameFlipHorizontal: chart.frameFlipHorizontal,
    frameFlipVertical: chart.frameFlipVertical,
  });
  return wrapAnchor(anchor, `${frame}${clientData}`);
}

export function stringifyWebExtension(
  webExtension: DrawingWebExtensionOptions,
  id: number,
  ctx?: WriteContext,
): string {
  const locks = webExtension.frameLocks
    ? (graphicFrameLockingDesc.stringify(webExtension.frameLocks, ctx as WriteContext) ?? "")
    : "";
  const cNvGraphicFramePr = locks
    ? `<xdr:cNvGraphicFramePr>${locks}</xdr:cNvGraphicFramePr>`
    : "<xdr:cNvGraphicFramePr/>";
  const objectAttrs = `${macroAttr(webExtension.macro)}${publishedObjectAttrs(webExtension)}`;
  const frame =
    `<xdr:graphicFrame${objectAttrs}><xdr:nvGraphicFramePr>${stringifyNonVisualDrawingProperties("xdr:cNvPr", id, webExtension, `WebExtension ${id}`)}` +
    `${cNvGraphicFramePr}</xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
    `<a:graphic><a:graphicData uri="${WE_URI}">` +
    `<we:${webExtension.elementName ?? "webextensionref"} xmlns:we="${WE_NS}" r:id="${webExtension.rId}"/>` +
    `</a:graphicData></a:graphic></xdr:graphicFrame>`;
  const fallback = webExtension.fallback
    ? `<mc:Fallback>${picXml(
        { col: 1, row: 1, ...webExtension.fallback },
        id,
        convertToEmu(webExtension.extentCx ?? DEFAULT_EXTENT_CX),
        convertToEmu(webExtension.extentCy ?? DEFAULT_EXTENT_CY),
        ctx as WriteContext,
      )}</mc:Fallback>`
    : "";
  const content =
    `<mc:AlternateContent xmlns:mc="${MC_NS}">` +
    `<mc:Choice xmlns:we="${WE_NS}" Requires="we">${frame}</mc:Choice>` +
    `${fallback}</mc:AlternateContent>`;
  return wrapAnchor(webExtension, `${content}${clientDataXml(webExtension)}`);
}

/** Anchored SmartArt: a graphicFrame whose graphicData points at the four
 *  diagram parts through dgm:relIds (same anchor footprint as a chart). */
export function stringifySmartArt(
  smartArt: DrawingSmartArtOptions,
  id: number,
  ctx?: WriteContext,
): string {
  const anchor = { toCol: smartArt.col + 9, toRow: smartArt.row + 16, ...smartArt };
  const clientData = clientDataXml(smartArt);
  const cellAnchored = (anchor.anchorType ?? ANCHOR_TYPES.twoCell) !== ANCHOR_TYPES.absolute;
  const cx = cellAnchored ? 0 : convertToEmu(anchor.extentCx ?? DEFAULT_EXTENT_CX);
  const cy = cellAnchored ? 0 : convertToEmu(anchor.extentCy ?? DEFAULT_EXTENT_CY);
  // Same nv/xfrm scaffolding as graphicFrameXml, with dgm:relIds replacing
  // the c:chart payload.
  const locks = smartArt.frameLocks
    ? (graphicFrameLockingDesc.stringify(smartArt.frameLocks, ctx as WriteContext) ?? "")
    : "";
  const cNvGraphicFramePr = locks
    ? `<xdr:cNvGraphicFramePr>${locks}</xdr:cNvGraphicFramePr>`
    : "<xdr:cNvGraphicFramePr/>";
  const objectAttrs = `${macroAttr(smartArt.macro)}${publishedObjectAttrs(smartArt)}`;
  const frame =
    `<xdr:graphicFrame${objectAttrs}><xdr:nvGraphicFramePr>${stringifyNonVisualDrawingProperties("xdr:cNvPr", id, smartArt, `SmartArt ${id}`)}` +
    `${cNvGraphicFramePr}</xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></xdr:xfrm>` +
    `<a:graphic><a:graphicData uri="${DGM_URI}">` +
    `<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="${smartArt.dataRId}" r:lo="${smartArt.layoutRId}" r:qs="${smartArt.quickStyleRId}" r:cs="${smartArt.colorsRId}"/>` +
    `</a:graphicData></a:graphic></xdr:graphicFrame>`;
  return wrapAnchor(anchor, `${frame}${clientData}`);
}

/** Build the inner xdr:sp content (nvSpPr + spPr + optional style/txBody). */
function buildShapeContent(
  shape: NonVisualDrawingPropertiesOptions & {
    blackWhiteMode?: BlackWhiteMode;
    textBox?: boolean;
    hyperlink?: TextHyperlinkOptions;
    locking?: ShapeLockingOptions;
    fPublished?: boolean;
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
  const txBoxAttr = shape.textBox === undefined ? "" : ` txBox="${shape.textBox ? 1 : 0}"`;
  const publishedAttr = publishedObjectAttrs(shape);
  const locksXml = shape.locking ? (shapeLockingDesc.stringify(shape.locking, ctx) ?? "") : "";
  const cNvSpPr = locksXml
    ? `<xdr:cNvSpPr${txBoxAttr}>${locksXml}</xdr:cNvSpPr>`
    : `<xdr:cNvSpPr${txBoxAttr}/>`;
  const bwModeAttr = shape.blackWhiteMode ? ` bwMode="${shape.blackWhiteMode}"` : "";
  return `<xdr:sp${attrs}${publishedAttr}><xdr:nvSpPr>${stringifyNonVisualDrawingProperties("xdr:cNvPr", id, cNvPr, fallbackName, hlinkClickXml(shape.hyperlink, ctx))}${cNvSpPr}</xdr:nvSpPr><xdr:spPr${bwModeAttr}>${spPrXml}</xdr:spPr>${styleXml}${txBodyXml}</xdr:sp>`;
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
    blackWhiteMode?: BlackWhiteMode;
    locking?: ConnectorLockingOptions;
    startConnection?: EndpointConnectionOptions;
    endConnection?: EndpointConnectionOptions;
    style?: DefaultShapeStyleOptions;
    fPublished?: boolean;
  },
): string {
  const bwModeAttr = connector?.blackWhiteMode ? ` bwMode="${connector.blackWhiteMode}"` : "";
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
  const publishedAttr = publishedObjectAttrs({ fPublished: connector?.fPublished });
  return `<xdr:cxnSp${attrs}${publishedAttr}><xdr:nvCxnSpPr>${stringifyNonVisualDrawingProperties("xdr:cNvPr", id, cNvPr, fallbackName, hlinkClickXml(cNvPr?.hyperlink, ctx))}${cNvCxnSpPr}</xdr:nvCxnSpPr><xdr:spPr${bwModeAttr}>${spPrXml}</xdr:spPr>${styleXml}</xdr:cxnSp>`;
}

export function stringifyShape(shape: ShapeOptions, id: number, ctx: WriteContext): string {
  const xml = buildShapeContent(
    shape,
    id,
    `Shape ${id}`,
    shape.properties,
    shape.textBody,
    ctx,
    macroTextlinkAttrs(shape),
    shape.style,
  );
  const content = shape.objectAlternateContent
    ? `<mc:AlternateContent xmlns:mc="${MC_NS}"><mc:Choice xmlns:a14="${A14_NS}" Requires="a14">${xml}</mc:Choice><mc:Fallback/></mc:AlternateContent>`
    : xml;
  return wrapAnchor(shape, `${content}${clientDataXml(shape)}`);
}

export function stringifyConnector(conn: ConnectorOptions, id: number, ctx: WriteContext): string {
  const xml = buildConnectorContent(
    conn,
    id,
    `Connector ${id}`,
    conn.properties,
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
  const grpSpPrXml = groupShapePropertiesDesc.stringify(grp.properties, ctx) ?? "";
  let nextChildId = id + 1;
  const children: string[] = [];
  for (const childImage of grp.images ?? []) {
    const childId = childImage.shapeId ?? nextChildId;
    nextChildId = childId + 1;
    children.push(
      picXml(
        childImage,
        childId,
        convertToEmu(childImage.extentCx ?? DEFAULT_EXTENT_CX),
        convertToEmu(childImage.extentCy ?? DEFAULT_EXTENT_CY),
        ctx,
      ),
    );
  }
  for (const childShape of grp.shapes ?? []) {
    const childId = childShape.shapeId ?? nextChildId;
    nextChildId = childId + 1;
    children.push(
      buildShapeContent(
        childShape,
        childId,
        `Shape ${childId}`,
        childShape.properties,
        childShape.textBody,
        ctx,
        macroTextlinkAttrs(childShape),
        childShape.style,
      ),
    );
  }
  for (const childConn of grp.connectors ?? []) {
    const childId = childConn.shapeId ?? nextChildId;
    nextChildId = childId + 1;
    children.push(
      buildConnectorContent(
        childConn,
        childId,
        `Connector ${childId}`,
        childConn.properties,
        ctx,
        macroTextlinkAttrs(childConn),
        childConn,
      ),
    );
  }
  const xml =
    `<xdr:grpSp><xdr:nvGrpSpPr>${stringifyNonVisualDrawingProperties("xdr:cNvPr", id, grp, `Group ${id}`, hlinkClickXml(grp.hyperlink, ctx))}` +
    (grp.locking === null
      ? "<xdr:cNvGrpSpPr><a:grpSpLocks/></xdr:cNvGrpSpPr>"
      : grp.locking
        ? `<xdr:cNvGrpSpPr>${groupLockingDesc.stringify(grp.locking, ctx) ?? ""}</xdr:cNvGrpSpPr>`
        : "<xdr:cNvGrpSpPr/>") +
    `</xdr:nvGrpSpPr>` +
    `<xdr:grpSpPr${grp.blackWhiteMode ? ` bwMode="${grp.blackWhiteMode}"` : ""}>${grpSpPrXml}</xdr:grpSpPr>${children.join("")}</xdr:grpSp>`;
  return { xml, nextId: nextChildId };
}

/** CT_Shape attribute string (macro/textlink) with leading space, or empty. */
function macroTextlinkAttrs(shape: {
  macro?: string;
  textlink?: string;
  fLocksText?: boolean;
}): string {
  const a: string[] = [];
  if (shape.macro !== undefined) a.push(`macro="${escapeXml(shape.macro)}"`);
  if (shape.textlink !== undefined) a.push(`textlink="${escapeXml(shape.textlink)}"`);
  if (shape.fLocksText !== undefined) a.push(`fLocksText="${shape.fLocksText ? 1 : 0}"`);
  return a.length ? " " + a.join(" ") : "";
}

function macroAttr(macro: string | undefined): string {
  return macro === undefined ? "" : ` macro="${escapeXml(macro)}"`;
}

function publishedObjectAttrs(source: { fPublished?: boolean }): string {
  return source.fPublished === undefined ? "" : ` fPublished="${source.fPublished ? 1 : 0}"`;
}
