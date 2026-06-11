/**
 * Drawing descriptor for DOCX documents.
 *
 * Produces `<w:drawing>` XML directly from media data and options,
 * eliminating the Drawing/Inline/Anchor/Graphic/GraphicData/Pic XmlComponent
 * class chain (~10 instances per drawing in the old path).
 *
 * Common path (inline image/chart/smartart without advanced properties):
 * zero XmlComponent instances — pure string concatenation.
 *
 * Advanced properties (outline, fill, effects on images) and floating
 * positioning use core `create*()` + `.toXml({stack:[]})` for sub-elements
 * (lightweight BuilderElement instances, not deep hierarchies).
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_Drawing
 *
 * @module
 */

import { TargetModeType } from "@office-open/core";
import { uniqueNumericIdCreator, uniqueId } from "@office-open/core";
import type { CustomDescriptor, WriteContext } from "@office-open/core/descriptor";
import type { FillOptions } from "@office-open/core/drawingml";
import {
  calculateEffectExtent,
  createEffectDag,
  customGeometryDesc,
  effectListDesc,
  extractBlipFillMedia,
  fillDesc,
  outlineDesc,
  scene3DDesc,
  shape3DDesc,
  transform2DDesc,
} from "@office-open/core/drawingml";
import { escapeXml } from "@office-open/xml";
import { stringifyParagraphInline } from "@parts/inline";
import type {
  ChartMediaData,
  IExtendedMediaData,
  IGroupChildMediaData,
  IMediaData,
  MediaDataTransformation,
  SmartArtMediaData,
  WpgMediaData,
  WpsMediaData,
} from "@shared/media";

import type { BodyContext } from "../../context";
import type { DocPropertiesOptions, HyperlinkOptions } from "./doc-properties/doc-properties";
// Import parse function from drawing-parse.ts (parse path)
import { parseDrawingRun } from "./drawing-parse";
import type { Floating, HorizontalPositionOptions, VerticalPositionOptions } from "./floating";
import type { Margins } from "./floating";
import { HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom } from "./floating";
import type { BlipEffectsOptions } from "./inline/graphic/graphic-data/pic/blip/blip-effects";
import type { TileOptions } from "./inline/graphic/graphic-data/pic/blip/tile";
import type { EffectListOptions } from "./inline/graphic/graphic-data/pic/effects/effect-list";
import type { OutlineOptions } from "./inline/graphic/graphic-data/pic/outline/outline";
import type { ChildOffset, ChildExtent } from "./inline/graphic/graphic-data/wpg/wpg-group";
// wpg/wps types only
import type { BodyPropertiesOptions } from "./inline/graphic/graphic-data/wps/body-properties";
import type { NonVisualShapePropertiesOptions } from "./inline/graphic/graphic-data/wps/non-visual-shape-properties";
import type { WpsShapeCoreOptions } from "./inline/graphic/graphic-data/wps/wps-shape";
import { TextWrappingSide, TextWrappingType } from "./text-wrap";
import type { TextWrapping } from "./text-wrap";

// Noop context for drawingml descriptors that don't use WriteContext
const NOOP_CTX: WriteContext = {
  addRelationship: () => "",
  addMedia: () => "",
};

// ── Options ──

/**
 * Options for the drawing descriptor.
 *
 * Combines media data with optional visual properties.
 */
export interface DrawingDescriptorOptions {
  /** Media data (image, chart, smartart, wps, wpg) */
  mediaData: IExtendedMediaData;
  /** Non-visual document properties (name, description, hyperlinks) */
  docProperties?: DocPropertiesOptions;
  /** Floating/anchored positioning (omit for inline) */
  floating?: Floating;
  /** Shape outline */
  outline?: OutlineOptions;
  /** Shape fill */
  fill?: FillOptions;
  /** Shape effects (shadow, glow, etc.) */
  effects?: EffectListOptions;
  /** Image blip effects (brightness, contrast, etc.) */
  blipEffects?: BlipEffectsOptions;
  /** Image tile fill mode */
  tile?: TileOptions;
}

// ── ID generation ──

let _docPropsIdGen = uniqueNumericIdCreator();

/** Reset the doc properties ID generator (for testing). */
export const resetDrawingIdGen = (): void => {
  _docPropsIdGen = uniqueNumericIdCreator();
};

// ── Constants ──

const GRAPHIC_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const PIC_URI = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const DGM_URI = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const WPS_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const WPG_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
const HYPERLINK_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

// ── Hyperlink handling ──

interface HyperlinkIds {
  clickId?: string;
  hoverId?: string;
}

function registerHyperlinks(
  hyperlink: HyperlinkOptions | undefined,
  ctx: BodyContext,
): HyperlinkIds {
  if (!hyperlink) return {};
  const result: HyperlinkIds = {};
  if (hyperlink.click) {
    const linkId = uniqueId();
    ctx.viewWrapper.relationships.addRelationship(
      linkId,
      HYPERLINK_REL,
      hyperlink.click,
      TargetModeType.EXTERNAL,
    );
    result.clickId = `rId${linkId}`;
  }
  if (hyperlink.hover) {
    const linkId = uniqueId();
    ctx.viewWrapper.relationships.addRelationship(
      linkId,
      HYPERLINK_REL,
      hyperlink.hover,
      TargetModeType.EXTERNAL,
    );
    result.hoverId = `rId${linkId}`;
  }
  return result;
}

function buildHyperlinkChildren(ids: HyperlinkIds): string {
  const parts: string[] = [];
  if (ids.clickId) parts.push(`<a:hlinkClick r:id="${ids.clickId}"/>`);
  if (ids.hoverId) parts.push(`<a:hlinkHover r:id="${ids.hoverId}"/>`);
  return parts.join("");
}

// ── DocPr ──

function stringifyDocPr(opts: DocPropertiesOptions | undefined, hlIds: HyperlinkIds): string {
  const id = opts?.id ?? _docPropsIdGen();
  const name = opts?.name ?? "";
  const attrs: string[] = [`id="${id}"`, `name="${escapeXml(name)}"`];
  if (opts?.description != null && opts.description !== undefined) {
    attrs.push(`descr="${escapeXml(opts.description)}"`);
  }
  if (opts?.title != null && opts.title !== undefined) {
    attrs.push(`title="${escapeXml(opts.title)}"`);
  }
  const hlXml = buildHyperlinkChildren(hlIds);
  if (hlXml) {
    return `<wp:docPr ${attrs.join(" ")}>${hlXml}</wp:docPr>`;
  }
  return `<wp:docPr ${attrs.join(" ")}/>`;
}

// ── BlipFill (image data reference) ──

function stringifyBlipFill(
  mediaData: IMediaData,
  blipEffects?: BlipEffectsOptions,
  tile?: TileOptions,
): string {
  const fileName =
    mediaData.type === "svg" && "fallback" in mediaData
      ? mediaData.fallback.fileName
      : mediaData.fileName;

  const parts: string[] = [];

  // a:blip
  const blipAttrs: string[] = [`r:embed="{${escapeXml(fileName)}}"`, 'cstate="none"'];

  // SVG extension list
  const svgExtXml =
    mediaData.type === "svg"
      ? `<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="{${escapeXml(mediaData.fileName)}}"/></a:ext></a:extLst>`
      : "";

  // Blip effects
  const blipEffectsXml = blipEffects ? buildBlipEffectsXml(blipEffects) : "";

  const blipContent = svgExtXml + blipEffectsXml;
  if (blipContent) {
    parts.push(`<a:blip ${blipAttrs.join(" ")}>${blipContent}</a:blip>`);
  } else {
    parts.push(`<a:blip ${blipAttrs.join(" ")}/>`);
  }

  // Source rectangle
  if (mediaData.srcRect) {
    const srAttrs: string[] = [];
    if (mediaData.srcRect.left !== undefined) srAttrs.push(`l="${mediaData.srcRect.left}"`);
    if (mediaData.srcRect.top !== undefined) srAttrs.push(`t="${mediaData.srcRect.top}"`);
    if (mediaData.srcRect.right !== undefined) srAttrs.push(`r="${mediaData.srcRect.right}"`);
    if (mediaData.srcRect.bottom !== undefined) srAttrs.push(`b="${mediaData.srcRect.bottom}"`);
    if (srAttrs.length) {
      parts.push(`<a:srcRect ${srAttrs.join(" ")}/>`);
    }
  }

  // Tile or stretch
  if (tile) {
    const tileAttrs: string[] = [];
    if (tile.tx !== undefined) tileAttrs.push(`tx="${tile.tx}"`);
    if (tile.ty !== undefined) tileAttrs.push(`ty="${tile.ty}"`);
    if (tile.sx !== undefined) tileAttrs.push(`sx="${tile.sx}"`);
    if (tile.sy !== undefined) tileAttrs.push(`sy="${tile.sy}"`);
    const tileAttrStr = tileAttrs.length ? " " + tileAttrs.join(" ") : "";
    parts.push(`<a:tile${tileAttrStr}/>`);
  } else {
    parts.push("<a:stretch><a:fillRect/></a:stretch>");
  }

  return `<pic:blipFill>${parts.join("")}</pic:blipFill>`;
}

function buildBlipEffectsXml(opts: BlipEffectsOptions): string {
  const parts: string[] = [];
  if (opts.grayscale) parts.push("<a:grayscl/>");
  if (opts.luminance) {
    const a: string[] = [];
    if (opts.luminance.bright !== undefined) a.push(`bright="${opts.luminance.bright}%"`);
    if (opts.luminance.contrast !== undefined) a.push(`contrast="${opts.luminance.contrast}%"`);
    parts.push(`<a:lum${a.length ? " " + a.join(" ") : ""}/>`);
  }
  if (opts.biLevel) parts.push(`<a:biLevel thresh="${opts.biLevel.threshold}%"/>`);
  if (opts.blur) {
    const a: string[] = [];
    if (opts.blur.radius !== undefined) a.push(`rad="${opts.blur.radius}"`);
    if (opts.blur.grow === false) a.push('grow="0"');
    parts.push(`<a:blur${a.length ? " " + a.join(" ") : ""}/>`);
  }
  return parts.join("");
}

// ── Shape Properties (pic:spPr) ──

function stringifyShapeProps(
  transform: MediaDataTransformation,
  outline?: OutlineOptions,
  fill?: FillOptions,
  effects?: EffectListOptions,
): string {
  const parts: string[] = [];

  // Transform
  parts.push(
    transform2DDesc.stringify(
      {
        x: transform.offset?.emus?.x ?? 0,
        y: transform.offset?.emus?.y ?? 0,
        width: transform.emus.x,
        height: transform.emus.y,
        flipHorizontal: transform.flip?.horizontal,
        flipVertical: transform.flip?.vertical,
        rotation: transform.rotation,
      },
      NOOP_CTX,
    ) ?? "",
  );

  // Geometry (always rect — preset geometry variations not used for pic)
  parts.push('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>');

  if (fill) parts.push(fillDesc.stringify(fill, NOOP_CTX) ?? "");
  if (outline) parts.push(outlineDesc.stringify(outline, NOOP_CTX) ?? "");
  if (effects) parts.push(effectListDesc.stringify(effects, NOOP_CTX) ?? "");

  return `<pic:spPr bwMode="auto">${parts.join("")}</pic:spPr>`;
}

// ── Non-visual picture properties (pic:nvPicPr) ──

function stringifyNvPicPr(hlIds: HyperlinkIds): string {
  const hlXml = buildHyperlinkChildren(hlIds);
  const cNvPrClose = hlXml ? `>${hlXml}</pic:cNvPr>` : "/>";
  return (
    `<pic:nvPicPr><pic:cNvPr id="0" name="" descr=""${cNvPrClose}` +
    `<pic:cNvPicPr preferRelativeResize="1"><a:picLocks noChangeArrowheads="1" noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr>`
  );
}

// ── WPS shape (pure string, no class instances) ──

function stringifyGroupTransform2D(
  transform: MediaDataTransformation,
  chOff?: ChildOffset,
  chExt?: ChildExtent,
): string {
  const attrs: string[] = [];
  if (transform.flip?.horizontal !== undefined) attrs.push(`flipH="${transform.flip.horizontal}"`);
  if (transform.flip?.vertical !== undefined) attrs.push(`flipV="${transform.flip.vertical}"`);
  if (transform.rotation !== undefined) attrs.push(`rot="${transform.rotation}"`);
  const attrStr = attrs.length ? " " + attrs.join(" ") : "";

  const off = `<a:off x="${transform.offset?.emus?.x ?? 0}" y="${transform.offset?.emus?.y ?? 0}"/>`;
  const ext = `<a:ext cx="${transform.emus.x}" cy="${transform.emus.y}"/>`;
  const chOffXml = chOff ? `<a:chOff x="${chOff.x}" y="${chOff.y}"/>` : "";
  const chExtXml = chExt ? `<a:chExt cx="${chExt.cx}" cy="${chExt.cy}"/>` : "";

  return `<a:xfrm${attrStr}>${off}${ext}${chOffXml}${chExtXml}</a:xfrm>`;
}

/** WpsShape options for stringification (extends WpsShapeCoreOptions with transformation). */
interface WpsStringifyOptions extends WpsShapeCoreOptions {
  transformation: MediaDataTransformation;
}

function stringifyWpsShape(opts: WpsStringifyOptions, ctx: BodyContext): string {
  const transform = opts.transformation;
  const spPrParts: string[] = [];
  spPrParts.push(
    transform2DDesc.stringify(
      {
        x: transform.offset?.emus?.x ?? 0,
        y: transform.offset?.emus?.y ?? 0,
        width: transform.emus.x,
        height: transform.emus.y,
        flipHorizontal: transform.flip?.horizontal,
        flipVertical: transform.flip?.vertical,
        rotation: transform.rotation,
      },
      NOOP_CTX,
    ) ?? "",
  );
  if (opts.customGeometry) {
    spPrParts.push(customGeometryDesc.stringify(opts.customGeometry, NOOP_CTX) ?? "");
  } else {
    spPrParts.push('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>');
  }
  if (opts.fill) spPrParts.push(fillDesc.stringify(opts.fill, NOOP_CTX) ?? "");
  if (opts.outline) spPrParts.push(outlineDesc.stringify(opts.outline, NOOP_CTX) ?? "");
  if (opts.effectDag) {
    spPrParts.push(createEffectDag(opts.effectDag));
  } else if (opts.effects) {
    spPrParts.push(effectListDesc.stringify(opts.effects, NOOP_CTX) ?? "");
  }
  if (opts.scene3d) spPrParts.push(scene3DDesc.stringify(opts.scene3d, NOOP_CTX) ?? "");
  if (opts.shape3d) spPrParts.push(shape3DDesc.stringify(opts.shape3d, NOOP_CTX) ?? "");

  // Non-visual shape properties — default txBox="1"
  const cNvSpPr = opts.nonVisualProperties
    ? stringifyNonVisualShapeProperties(opts.nonVisualProperties)
    : '<wps:cNvSpPr txBox="1"/>';

  // Paragraph children — pure JSON stringification
  const childXml =
    opts.children
      ?.map((c) =>
        stringifyParagraphInline(
          c as import("@parts/paragraph/paragraph").ParagraphOptions | string,
          ctx,
        ),
      )
      .join("") ?? "";

  return (
    "<wps:wsp>" +
    cNvSpPr +
    `<wps:spPr bwMode="auto">${spPrParts.join("")}</wps:spPr>` +
    `<wps:txbx><wps:txbxContent>${childXml}</wps:txbxContent></wps:txbx>` +
    stringifyBodyPr(opts.bodyProperties) +
    "</wps:wsp>"
  );
}

function stringifyNonVisualShapeProperties(opts: NonVisualShapePropertiesOptions): string {
  if (!opts.txBox) return "<wps:cNvSpPr/>";
  return `<wps:cNvSpPr txBox="${opts.txBox}"/>`;
}

function stringifyBodyPr(opts?: BodyPropertiesOptions): string {
  if (!opts) return "<wps:bodyPr/>";
  const attrs: string[] = [];
  if (opts.anchor !== undefined) attrs.push(`anchor="${opts.anchor}"`);
  if (opts.vert !== undefined) attrs.push(`vert="${opts.vert}"`);
  if (opts.wrap !== undefined) attrs.push(`wrap="${opts.wrap}"`);
  if (opts.lIns !== undefined) attrs.push(`lIns="${opts.lIns}"`);
  if (opts.tIns !== undefined) attrs.push(`tIns="${opts.tIns}"`);
  if (opts.rIns !== undefined) attrs.push(`rIns="${opts.rIns}"`);
  if (opts.bIns !== undefined) attrs.push(`bIns="${opts.bIns}"`);
  if (opts.numCol !== undefined) attrs.push(`numCol="${opts.numCol}"`);
  if (opts.rotation !== undefined) attrs.push(`rot="${opts.rotation}"`);
  const attrStr = attrs.length ? " " + attrs.join(" ") : "";

  // Sub-elements
  const parts: string[] = [];
  if (opts.normAutofit) {
    const afAttrs: string[] = [];
    if (opts.normAutofit.fontScale !== undefined)
      afAttrs.push(`fontScale="${opts.normAutofit.fontScale}"`);
    if (opts.normAutofit.lnSpcReduction !== undefined)
      afAttrs.push(`lnSpcReduction="${opts.normAutofit.lnSpcReduction}"`);
    parts.push(`<a:normAutofit ${afAttrs.join(" ")}/>`);
  }
  const body = parts.join("");
  return body ? `<wps:bodyPr${attrStr}>${body}</wps:bodyPr>` : `<wps:bodyPr${attrStr}/>`;
}

// ── WPG group (pure string, no class instances) ──

function stringifyWpgGroup(
  opts: {
    readonly children: readonly IGroupChildMediaData[];
    readonly transformation: MediaDataTransformation;
    readonly chOff?: ChildOffset;
    readonly chExt?: ChildExtent;
    readonly fill?: FillOptions;
    readonly effects?: EffectListOptions;
  },
  ctx: BodyContext,
): string {
  const transform = opts.transformation;
  const grpSpPrParts: string[] = [];
  grpSpPrParts.push(stringifyGroupTransform2D(transform, opts.chOff, opts.chExt));
  if (opts.fill) grpSpPrParts.push(fillDesc.stringify(opts.fill, NOOP_CTX) ?? "");
  if (opts.effects) grpSpPrParts.push(effectListDesc.stringify(opts.effects, NOOP_CTX) ?? "");

  // Children — wps shapes or pic elements
  const childXml = opts.children
    .map((child) => {
      if (child.type === "wps") {
        const wpsData = child as WpsMediaData & { outline?: OutlineOptions; fill?: FillOptions };
        return stringifyWpsShape(
          {
            ...wpsData.data,
            outline: wpsData.outline ?? wpsData.data.outline,
            fill: wpsData.fill ?? wpsData.data.fill,
            transformation: wpsData.transformation,
          },
          ctx,
        );
      }
      // pic child (IMediaData)
      const picData = child as IMediaData;
      const picParts: string[] = [];
      picParts.push(
        '<pic:nvPicPr><pic:cNvPr id="0" name="" descr=""/><pic:cNvPicPr preferRelativeResize="1"><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr>',
      );
      picParts.push(
        `<pic:blipFill><a:blip r:embed="{${picData.fileName}}" cstate="none"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`,
      );
      picParts.push(
        `<pic:spPr bwMode="auto">${
          transform2DDesc.stringify(
            {
              x: picData.transformation.offset?.emus?.x ?? 0,
              y: picData.transformation.offset?.emus?.y ?? 0,
              width: picData.transformation.emus.x,
              height: picData.transformation.emus.y,
            },
            NOOP_CTX,
          ) ?? "<a:xfrm/>"
        }<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`,
      );
      return `<pic:pic xmlns:pic="${PIC_URI}">${picParts.join("")}</pic:pic>`;
    })
    .join("");

  return (
    "<wpg:wgp>" +
    "<wpg:cNvGrpSpPr/>" +
    `<wpg:grpSpPr>${grpSpPrParts.join("")}</wpg:grpSpPr>` +
    childXml +
    "</wpg:wgp>"
  );
}

// ── Graphic data content ──

function stringifyGraphicDataContent(
  mediaData: IExtendedMediaData,
  opts: DrawingDescriptorOptions,
  hlIds: HyperlinkIds,
  ctx: BodyContext,
): string {
  const { outline, fill, effects, blipEffects, tile } = opts;
  const transform = mediaData.transformation;

  if (mediaData.type === "chart") {
    const md = mediaData as ChartMediaData;
    return (
      `<a:graphicData uri="${CHART_URI}">` +
      `<c:chart xmlns:c="${CHART_URI}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{chart:${md.chartKey}}"/>` +
      `</a:graphicData>`
    );
  }

  if (mediaData.type === "smartart") {
    const md = mediaData as SmartArtMediaData;
    return (
      `<a:graphicData uri="${DGM_URI}">` +
      `<dgm:relIds xmlns:dgm="${DGM_URI}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:dm="{smartart:${md.smartArtKey}}" r:lo="{smartart-lo:${md.smartArtKey}}" r:qs="{smartart-qs:${md.smartArtKey}}" r:cs="{smartart-cs:${md.smartArtKey}}"/>` +
      `</a:graphicData>`
    );
  }

  if (mediaData.type === "wps") {
    const md = mediaData as WpsMediaData;
    const wpsXml = stringifyWpsShape(
      {
        ...md.data,
        outline,
        fill,
        transformation: transform,
      },
      ctx,
    );
    return `<a:graphicData uri="${WPS_URI}">${wpsXml}</a:graphicData>`;
  }

  if (mediaData.type === "wpg") {
    const md = mediaData as WpgMediaData;
    const wpgXml = stringifyWpgGroup(
      {
        children: md.children,
        transformation: transform,
        chOff: md.chOff,
        chExt: md.chExt,
        fill: md.fill,
        effects: md.effects,
      },
      ctx,
    );
    return `<a:graphicData uri="${WPG_URI}">${wpgXml}</a:graphicData>`;
  }

  // Default: image (pic:pic)
  const md = mediaData as IMediaData;
  return (
    `<a:graphicData uri="${PIC_URI}">` +
    `<pic:pic xmlns:pic="${PIC_URI}">` +
    stringifyNvPicPr(hlIds) +
    stringifyBlipFill(md, blipEffects, tile) +
    stringifyShapeProps(transform, outline, fill, effects) +
    `</pic:pic></a:graphicData>`
  );
}

// ── Position helpers (for anchor) ──

function stringifyPositionH(opts: HorizontalPositionOptions): string {
  const rel = opts.relative ?? HorizontalPositionRelativeFrom.PAGE;
  const child = opts.align
    ? `<wp:align>${opts.align}</wp:align>`
    : opts.offset !== undefined
      ? `<wp:posOffset>${opts.offset}</wp:posOffset>`
      : "<wp:align>left</wp:align>";
  return `<wp:positionH relativeFrom="${rel}">${child}</wp:positionH>`;
}

function stringifyPositionV(opts: VerticalPositionOptions): string {
  const rel = opts.relative ?? VerticalPositionRelativeFrom.PAGE;
  const child = opts.align
    ? `<wp:align>${opts.align}</wp:align>`
    : opts.offset !== undefined
      ? `<wp:posOffset>${opts.offset}</wp:posOffset>`
      : "<wp:align>top</wp:align>";
  return `<wp:positionV relativeFrom="${rel}">${child}</wp:positionV>`;
}

// ── Text wrapping string builders ──

function wrapPolygonStr(cx: number, cy: number): string {
  return (
    `<wp:wrapPolygon edited="0">` +
    `<wp:start x="0" y="0"/>` +
    `<wp:lineTo x="0" y="${-cy}"/>` +
    `<wp:lineTo x="${cx}" y="${-cy}"/>` +
    `<wp:lineTo x="${cx}" y="0"/>` +
    `<wp:lineTo x="0" y="0"/>` +
    `</wp:wrapPolygon>`
  );
}

function wrapSquareStr(textWrapping: TextWrapping, margins?: Margins): string {
  const side = textWrapping.side ?? TextWrappingSide.BOTH_SIDES;
  const m = margins ?? {};
  const a = [
    `wrapText="${side}"`,
    ...(m.top != null ? [`distT="${m.top}"`] : []),
    ...(m.bottom != null ? [`distB="${m.bottom}"`] : []),
    ...(m.left != null ? [`distL="${m.left}"`] : []),
    ...(m.right != null ? [`distR="${m.right}"`] : []),
  ].join(" ");
  return `<wp:wrapSquare ${a}/>`;
}

function wrapTightStr(
  textWrapping: TextWrapping,
  margins: Margins,
  cx: number,
  cy: number,
): string {
  const side = textWrapping.side ?? TextWrappingSide.BOTH_SIDES;
  const a = [`wrapText="${side}"`];
  if (margins.left != null) a.push(`distL="${margins.left}"`);
  if (margins.right != null) a.push(`distR="${margins.right}"`);
  return `<wp:wrapTight ${a.join(" ")}>${wrapPolygonStr(cx, cy)}</wp:wrapTight>`;
}

function wrapThroughStr(
  textWrapping: TextWrapping,
  margins: Margins,
  cx: number,
  cy: number,
): string {
  const side = textWrapping.side ?? TextWrappingSide.BOTH_SIDES;
  const a = [`wrapText="${side}"`];
  if (margins.left != null) a.push(`distL="${margins.left}"`);
  if (margins.right != null) a.push(`distR="${margins.right}"`);
  return `<wp:wrapThrough ${a.join(" ")}>${wrapPolygonStr(cx, cy)}</wp:wrapThrough>`;
}

function wrapTopAndBottomStr(margins?: Margins): string {
  const m = margins ?? {};
  const a = [
    ...(m.top != null ? [`distT="${m.top}"`] : []),
    ...(m.bottom != null ? [`distB="${m.bottom}"`] : []),
  ].join(" ");
  return a ? `<wp:wrapTopAndBottom ${a}/>` : "<wp:wrapTopAndBottom/>";
}

// ── Inline wrapper ──

function stringifyInline(
  opts: DrawingDescriptorOptions,
  hlIds: HyperlinkIds,
  ctx: BodyContext,
): string {
  const { mediaData, effects, docProperties } = opts;
  const cx = mediaData.transformation.emus.x;
  const cy = mediaData.transformation.emus.y;

  // Calculate effectExtent based on actual effects (shadow, glow, reflection, softEdge)
  const effectExtent = calculateEffectExtent(effects);
  const graphicDataXml = stringifyGraphicDataContent(mediaData, opts, hlIds, ctx);

  return (
    `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:effectExtent l="${effectExtent.l}" t="${effectExtent.t}" r="${effectExtent.r}" b="${effectExtent.b}"/>` +
    stringifyDocPr(docProperties, hlIds) +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic ${GRAPHIC_NS}>${graphicDataXml}</a:graphic>` +
    `</wp:inline></w:drawing>`
  );
}

// ── Anchor (floating) wrapper ──

function stringifyAnchor(
  opts: DrawingDescriptorOptions,
  hlIds: HyperlinkIds,
  ctx: BodyContext,
): string {
  const { mediaData, floating: rawFloating, docProperties } = opts;
  const cx = mediaData.transformation.emus.x;
  const cy = mediaData.transformation.emus.y;

  const floating: Required<Floating> = {
    allowOverlap: true,
    behindDocument: false,
    horizontalPosition: {},
    layoutInCell: true,
    lockAnchor: false,
    verticalPosition: {},
    zIndex: mediaData.transformation.emus.y,
    margins: {},
    wrap: { type: TextWrappingType.NONE },
    ...rawFloating,
  };

  const attrParts = [
    `distT="${floating.margins?.top ?? 0}"`,
    `distB="${floating.margins?.bottom ?? 0}"`,
    `distL="${floating.margins?.left ?? 0}"`,
    `distR="${floating.margins?.right ?? 0}"`,
    'simplePos="0"',
    `allowOverlap="${floating.allowOverlap ? 1 : 0}"`,
    `behindDoc="${floating.behindDocument ? 1 : 0}"`,
    `locked="${floating.lockAnchor ? 1 : 0}"`,
    `layoutInCell="${floating.layoutInCell ? 1 : 0}"`,
    `relativeHeight="${floating.zIndex}"`,
  ];

  // Wrap
  let wrapXml: string;
  const rawWrap = rawFloating?.wrap;
  if (rawWrap?.type === TextWrappingType.SQUARE) {
    wrapXml = wrapSquareStr(rawWrap, floating.margins);
  } else if (rawWrap?.type === TextWrappingType.TIGHT) {
    wrapXml = wrapTightStr(rawWrap, floating.margins, cx, cy);
  } else if (rawWrap?.type === TextWrappingType.THROUGH) {
    wrapXml = wrapThroughStr(rawWrap, floating.margins, cx, cy);
  } else if (rawWrap?.type === TextWrappingType.TOP_AND_BOTTOM) {
    wrapXml = wrapTopAndBottomStr(floating.margins);
  } else {
    wrapXml = "<wp:wrapNone/>";
  }

  const graphicDataXml = stringifyGraphicDataContent(mediaData, opts, hlIds, ctx);

  return (
    `<w:drawing><wp:anchor ${attrParts.join(" ")}>` +
    '<wp:simplePos x="0" y="0"/>' +
    stringifyPositionH(floating.horizontalPosition) +
    stringifyPositionV(floating.verticalPosition) +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    wrapXml +
    stringifyDocPr(docProperties, hlIds) +
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
    `<a:graphic ${GRAPHIC_NS}>${graphicDataXml}</a:graphic>` +
    `</wp:anchor></w:drawing>`
  );
}

// ── Descriptor ──

/**
 * Drawing descriptor for DOCX `<w:drawing>` elements.
 *
 * Eliminates the Drawing/Inline/Anchor/Graphic/GraphicData/Pic XmlComponent
 * class chain. Inline images, charts, and smartarts produce XML via pure
 * string concatenation — zero XmlComponent instances.
 *
 * @example
 * ```typescript
 * const xml = drawingDesc.stringify({ mediaData, docProperties: opts.altText, floating: opts.floating }, ctx);
 * ```
 */
export const drawingDesc: CustomDescriptor<DrawingDescriptorOptions, BodyContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    // Register blip fill media from fill options (shape fill with image)
    if (opts.fill) {
      const media = extractBlipFillMedia(opts.fill);
      if (media) {
        ctx.file.media.addImage(media.fileName, {
          data: media.data,
          fileName: media.fileName,
          type: media.type as "png",
          transformation: { pixels: { x: 0, y: 0 }, emus: { x: 0, y: 0 } },
        });
      }
    }

    // Register hyperlink relationships
    const hlIds = registerHyperlinks(opts.docProperties?.hyperlink, ctx);

    if (opts.floating) {
      return stringifyAnchor(opts, hlIds, ctx);
    }
    return stringifyInline(opts, hlIds, ctx);
  },

  parse(el, ctx) {
    const result = parseDrawingRun(el, ctx as import("../../context").DocxReadContext);
    return result ?? {};
  },
};
