/**
 * Chart user shapes part (cdr:userShapes / CT_Drawing in dml-chartDrawing).
 *
 * Freeform shapes layered over a chart, anchored relatively (from/to markers,
 * 0.0–1.0 of the chart area) or absolutely (from marker + EMU extent). The
 * part hangs off the chart part via `c:userShapes r:id` — that reference is
 * produced and parsed by chartSpaceDesc; this module is the part body.
 *
 * The shape payload reuses the core drawingml descriptors: shapeProperties /
 * textBody / blipFill / groupShapeProperties / style-matrix / cNvPr.
 *
 * Reference: ISO/IEC 29500-4, dml-chartDrawing.xsd, CT_Drawing
 *
 * @module
 */

import { escapeXml, findChild, stringifyElement } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type { CustomDescriptor, ReadContext, WriteContext } from "../../descriptor";
import { stringify, parse } from "../../descriptor";
import { blipFillDesc } from "../../drawing/blip/blip-descriptors";
import type { BlipFillOptions } from "../../drawing/blip/blip-fill";
import { groupShapePropertiesDesc } from "../../drawing/group-shape-properties-desc";
import type { GroupShapePropertiesOptions } from "../../drawing/group-shape-properties-desc";
import type { GraphicFrameLockingOptions } from "../../drawing/locking/locking";
import { stringifyNonVisualDrawingProperties } from "../../drawing/non-visual";
import type { NonVisualDrawingPropertiesOptions } from "../../drawing/non-visual";
import { shapePropertiesDesc } from "../../drawing/shape-properties-desc";
import type { ShapePropertiesOptions } from "../../drawing/shape-properties-desc";
import type { TextBodyOptions } from "../../drawing/text/text-body";
import { textBodyDesc } from "../../drawing/text/text-body";
import type { Transform2DOptions } from "../../drawing/transform";
import { transform2DDesc } from "../../drawing/transform-descriptors";
import { stringifyShapeStyle, parseShapeStyle } from "../../theme/style-matrix";
import type { DefaultShapeStyleOptions } from "../../theme/theme-options";
import { parseOnOff } from "../../util/values";

const CDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

// Noop context for descriptors that don't touch relationships (style matrix
// colors never register media here; blip fills reference via referenceId).
const DIRECT_CTX: WriteContext = {
  addRelationship: () => "",
  addMedia: () => "",
  addHyperlink: () => {},
};

// ── Anchors ──

/** Anchor marker (CT_Marker): x/y as 0.0–1.0 fractions of the chart area. */
export interface UserShapeMarkerOptions {
  x: number;
  y: number;
}

/** Anchor extent (a:ext): width/height in EMU. */
export interface UserShapeExtentOptions {
  width: number;
  height: number;
}

/** Relative anchor (cdr:relSizeAnchor): the object spans from `from` to `to`. */
export interface RelativeSizeAnchorOptions {
  from: UserShapeMarkerOptions;
  to: UserShapeMarkerOptions;
  object: UserShapeObjectOptions;
}

/** Absolute anchor (cdr:absSizeAnchor): fixed EMU size at the `from` marker. */
export interface AbsoluteSizeAnchorOptions {
  from: UserShapeMarkerOptions;
  extent: UserShapeExtentOptions;
  object: UserShapeObjectOptions;
}

// ── Object choices (EG_ObjectChoices) ──

/** Fields shared by every user-shape object (cNvPr @id + @macro/@fPublished). */
interface ObjectCommonAttributes {
  /** cNvPr drawing id (schema-required) */
  id: number;
  /** Macro reference (@macro) */
  macro?: string;
  /** Published as a server-side shape (@fPublished, default false) */
  published?: boolean;
}

export interface UserShapeShapeOptions extends ObjectCommonAttributes {
  type: "shape";
  /** cNvPr — id is required by the schema */
  nonVisualProperties?: NonVisualDrawingPropertiesOptions;
  /** Text-box marker (cdr:cNvSpPr @txBox) */
  textBox?: boolean;
  shapeProperties: ShapePropertiesOptions;
  style?: DefaultShapeStyleOptions;
  textBody?: TextBodyOptions;
  /** Formula linking the text to a cell (@textlink) */
  textLink?: string;
  /** Text locking (@fLocksText, default true) */
  locksText?: boolean;
}

export interface UserShapeConnectorOptions extends ObjectCommonAttributes {
  type: "connector";
  nonVisualProperties?: NonVisualDrawingPropertiesOptions;
  shapeProperties: ShapePropertiesOptions;
  style?: DefaultShapeStyleOptions;
}

export interface UserShapePictureOptions extends ObjectCommonAttributes {
  type: "picture";
  nonVisualProperties?: NonVisualDrawingPropertiesOptions;
  /** Image reference key (a:blip r:embed via the relationship placeholder) */
  referenceId: string;
  blipFill?: BlipFillOptions;
  shapeProperties: ShapePropertiesOptions;
  style?: DefaultShapeStyleOptions;
}

export interface UserShapeGraphicFrameOptions extends ObjectCommonAttributes {
  type: "graphicFrame";
  nonVisualProperties?: NonVisualDrawingPropertiesOptions;
  /** Frame locking flags (cdr:cNvGraphicFramePr) */
  graphicFrameLocks?: GraphicFrameLockingOptions | null;
  /** Placement transform (cdr:xfrm) */
  transform: Transform2DOptions;
  /**
   * Verbatim a:graphic XML payload. The graphic contents (chart/table
   * references) are package-wired; the part body carries them unchanged.
   */
  graphic: string;
}

export interface UserShapeGroupOptions extends ObjectCommonAttributes {
  type: "group";
  nonVisualProperties?: NonVisualDrawingPropertiesOptions;
  groupShapeProperties: GroupShapePropertiesOptions;
  children: UserShapeObjectOptions[];
}

export type UserShapeObjectOptions =
  | UserShapeShapeOptions
  | UserShapeConnectorOptions
  | UserShapePictureOptions
  | UserShapeGraphicFrameOptions
  | UserShapeGroupOptions;

/** cdr:userShapes part options (CT_Drawing). */
export interface UserShapesOptions {
  anchors: (RelativeSizeAnchorOptions | AbsoluteSizeAnchorOptions)[];
}

// ── Stringify ──

function stringifyMarker(marker: UserShapeMarkerOptions): string {
  return `<cdr:x>${marker.x}</cdr:x><cdr:y>${marker.y}</cdr:y>`;
}

function stringifyCnvPr(id: number, nvp: NonVisualDrawingPropertiesOptions | undefined): string {
  return stringifyNonVisualDrawingProperties("cdr:cNvPr", id, nvp, "User Shape");
}

function commonAttrs(opts: ObjectCommonAttributes): string {
  let attrs = "";
  if (opts.macro !== undefined) attrs += ` macro="${escapeXml(opts.macro)}"`;
  if (opts.published) attrs += ' fPublished="1"';
  return attrs;
}

function isRelative(
  anchor: RelativeSizeAnchorOptions | AbsoluteSizeAnchorOptions,
): anchor is RelativeSizeAnchorOptions {
  return "to" in anchor;
}

function stringifyObject(obj: UserShapeObjectOptions): string {
  switch (obj.type) {
    case "shape": {
      const textLinkAttr =
        obj.textLink !== undefined ? ` textlink="${escapeXml(obj.textLink)}"` : "";
      const locksTextAttr =
        obj.locksText !== undefined ? ` fLocksText="${obj.locksText ? 1 : 0}"` : "";
      const nvSpPr =
        `<cdr:nvSpPr>${stringifyCnvPr(obj.id, obj.nonVisualProperties)}` +
        `<cdr:cNvSpPr${obj.textBox ? ' txBox="1"' : ""}/></cdr:nvSpPr>`;
      const spPrXml = stringify(shapePropertiesDesc, obj.shapeProperties, DIRECT_CTX) ?? "";
      const styleXml = obj.style ? stringifyShapeStyle(obj.style, DIRECT_CTX) : "";
      const txBodyXml = obj.textBody
        ? (stringify(textBodyDesc, obj.textBody, DIRECT_CTX) ?? "")
        : "";
      return (
        `<cdr:sp${commonAttrs(obj)}${textLinkAttr}${locksTextAttr}>` +
        nvSpPr +
        `<cdr:spPr>${spPrXml}</cdr:spPr>` +
        styleXml +
        txBodyXml +
        "</cdr:sp>"
      );
    }
    case "connector": {
      const nvCxnSpPr = `<cdr:nvCxnSpPr>${stringifyCnvPr(obj.id, obj.nonVisualProperties)}<cdr:cNvCxnSpPr/></cdr:nvCxnSpPr>`;
      const spPrXml = stringify(shapePropertiesDesc, obj.shapeProperties, DIRECT_CTX) ?? "";
      const styleXml = obj.style ? stringifyShapeStyle(obj.style, DIRECT_CTX) : "";
      return (
        `<cdr:cxnSp${commonAttrs(obj)}>` +
        nvCxnSpPr +
        `<cdr:spPr>${spPrXml}</cdr:spPr>` +
        styleXml +
        "</cdr:cxnSp>"
      );
    }
    case "picture": {
      const nvPicPr = `<cdr:nvPicPr>${stringifyCnvPr(obj.id, obj.nonVisualProperties)}<cdr:cNvPicPr/></cdr:nvPicPr>`;
      const fillOpts = { ...obj.blipFill, referenceId: obj.referenceId };
      // blipFill is a local element of CT_Picture → cdr-prefixed root; the
      // blip/srcRect/tile children belong to a: either way.
      const blipFillXml = (stringify(blipFillDesc, fillOpts, DIRECT_CTX) ?? "").replace(
        /pic:blipFill/g,
        "cdr:blipFill",
      );
      const spPrXml = stringify(shapePropertiesDesc, obj.shapeProperties, DIRECT_CTX) ?? "";
      const styleXml = obj.style ? stringifyShapeStyle(obj.style, DIRECT_CTX) : "";
      return (
        `<cdr:pic${commonAttrs(obj)}>` +
        nvPicPr +
        blipFillXml +
        `<cdr:spPr>${spPrXml}</cdr:spPr>` +
        styleXml +
        "</cdr:pic>"
      );
    }
    case "graphicFrame": {
      const locks = obj.graphicFrameLocks;
      const locksXml = locks
        ? `<a:graphicFrameLocks${Object.entries(locks)
            .filter(([, v]) => v)
            .map(([k]) => ` ${k}="1"`)
            .join("")}/>`
        : "";
      const nvGraphicFramePr = `<cdr:nvGraphicFramePr>${stringifyCnvPr(
        obj.id,
        obj.nonVisualProperties,
      )}<cdr:cNvGraphicFramePr>${locksXml}</cdr:cNvGraphicFramePr></cdr:nvGraphicFramePr>`;
      // xfrm is a local element of CT_GraphicFrame → cdr-prefixed root; the
      // off/ext children belong to a: either way, so re-tag the root only.
      const xfrmXml = (stringify(transform2DDesc, obj.transform, DIRECT_CTX) ?? "").replace(
        /a:xfrm/g,
        "cdr:xfrm",
      );
      return (
        `<cdr:graphicFrame${commonAttrs(obj)}>` +
        nvGraphicFramePr +
        xfrmXml +
        obj.graphic +
        "</cdr:graphicFrame>"
      );
    }
    case "group": {
      const nvGrpSpPr = `<cdr:nvGrpSpPr>${stringifyCnvPr(obj.id, obj.nonVisualProperties)}<cdr:cNvGrpSpPr/></cdr:nvGrpSpPr>`;
      const grpSpPrXml =
        stringify(groupShapePropertiesDesc, obj.groupShapeProperties, DIRECT_CTX) ?? "";
      return (
        "<cdr:grpSp>" +
        nvGrpSpPr +
        `<cdr:grpSpPr>${grpSpPrXml}</cdr:grpSpPr>` +
        obj.children.map(stringifyObject).join("") +
        "</cdr:grpSp>"
      );
    }
  }
}

// ── Parse ──

function readMarker(el: XmlElement | undefined): UserShapeMarkerOptions | undefined {
  if (!el) return undefined;
  const x = findChild(el, "cdr:x");
  const y = findChild(el, "cdr:y");
  if (!x || !y) return undefined;
  return {
    x: Number(x.elements?.[0]?.text ?? 0),
    y: Number(y.elements?.[0]?.text ?? 0),
  };
}

interface ParsedCnvPr {
  id: number;
  nvp: NonVisualDrawingPropertiesOptions;
}

function readCnvPr(el: XmlElement | undefined): ParsedCnvPr | undefined {
  if (!el) return undefined;
  const attrs = el.attributes ?? {};
  if (attrs["id"] === undefined) return undefined;
  const nvp: Partial<NonVisualDrawingPropertiesOptions> = {};
  if (attrs["name"] !== undefined) nvp.name = String(attrs["name"]);
  if (attrs["descr"] !== undefined) nvp.description = String(attrs["descr"]);
  if (attrs["hidden"] !== undefined) nvp.hidden = parseOnOff(attrs["hidden"]) ?? false;
  return { id: Number(attrs["id"]), nvp: nvp as NonVisualDrawingPropertiesOptions };
}

/** macro/fPublished attributes shared by every non-group user-shape object. */
function readCommonAttrs(el: XmlElement): { macro?: string; published?: boolean } {
  const out: { macro?: string; published?: boolean } = {};
  if (el.attributes?.["macro"] !== undefined) out.macro = String(el.attributes["macro"]);
  if (el.attributes?.["fPublished"] !== undefined)
    out.published = parseOnOff(el.attributes["fPublished"]) ?? false;
  return out;
}

function readObject(el: XmlElement, ctx: ReadContext): UserShapeObjectOptions | undefined {
  switch (el.name) {
    case "cdr:sp": {
      const nvSpPr = findChild(el, "cdr:nvSpPr");
      const spPr = findChild(el, "cdr:spPr");
      if (!nvSpPr || !spPr) return undefined;
      const cnvPr = readCnvPr(findChild(nvSpPr, "cdr:cNvPr"));
      if (!cnvPr) return undefined;
      const cNvSpPr = findChild(nvSpPr, "cdr:cNvSpPr");
      const result: UserShapeShapeOptions = {
        type: "shape",
        id: cnvPr.id,
        nonVisualProperties: cnvPr.nvp,
        shapeProperties: parse(shapePropertiesDesc, spPr, ctx),
      };
      if (cNvSpPr && cNvSpPr.attributes?.["txBox"] !== undefined)
        result.textBox = parseOnOff(cNvSpPr.attributes["txBox"]) ?? false;
      const styleEl = findChild(el, "a:style");
      if (styleEl) result.style = parseShapeStyle(styleEl, ctx);
      const txBody = findChild(el, "cdr:txBody");
      if (txBody) result.textBody = parse(textBodyDesc, txBody, ctx);
      Object.assign(result, readCommonAttrs(el));
      if (el.attributes?.["textlink"] !== undefined)
        result.textLink = String(el.attributes["textlink"]);
      if (el.attributes?.["fLocksText"] !== undefined)
        result.locksText = parseOnOff(el.attributes["fLocksText"]) ?? true;
      return result;
    }
    case "cdr:cxnSp": {
      const nvCxnSpPr = findChild(el, "cdr:nvCxnSpPr");
      const spPr = findChild(el, "cdr:spPr");
      if (!nvCxnSpPr || !spPr) return undefined;
      const cnvPr = readCnvPr(findChild(nvCxnSpPr, "cdr:cNvPr"));
      if (!cnvPr) return undefined;
      const result: UserShapeConnectorOptions = {
        type: "connector",
        id: cnvPr.id,
        nonVisualProperties: cnvPr.nvp,
        shapeProperties: parse(shapePropertiesDesc, spPr, ctx),
      };
      const styleEl = findChild(el, "a:style");
      if (styleEl) result.style = parseShapeStyle(styleEl, ctx);
      Object.assign(result, readCommonAttrs(el));
      return result;
    }
    case "cdr:pic": {
      const nvPicPr = findChild(el, "cdr:nvPicPr");
      const blipFill = findChild(el, "cdr:blipFill");
      const spPr = findChild(el, "cdr:spPr");
      if (!nvPicPr || !blipFill || !spPr) return undefined;
      const cnvPr = readCnvPr(findChild(nvPicPr, "cdr:cNvPr"));
      if (!cnvPr) return undefined;
      const fill = parse(blipFillDesc, blipFill, ctx);
      const { referenceId, ...blipFillOpts } = fill;
      const result: UserShapePictureOptions = {
        type: "picture",
        id: cnvPr.id,
        nonVisualProperties: cnvPr.nvp,
        referenceId: referenceId ?? "",
        shapeProperties: parse(shapePropertiesDesc, spPr, ctx),
      };
      if (Object.keys(blipFillOpts).length > 0) result.blipFill = blipFillOpts;
      const styleEl = findChild(el, "a:style");
      if (styleEl) result.style = parseShapeStyle(styleEl, ctx);
      Object.assign(result, readCommonAttrs(el));
      return result;
    }
    case "cdr:graphicFrame": {
      const nvGraphicFramePr = findChild(el, "cdr:nvGraphicFramePr");
      const xfrm = findChild(el, "cdr:xfrm");
      const graphic = findChild(el, "a:graphic");
      if (!nvGraphicFramePr || !xfrm || !graphic) return undefined;
      const cnvPr = readCnvPr(findChild(nvGraphicFramePr, "cdr:cNvPr"));
      if (!cnvPr) return undefined;
      const result: UserShapeGraphicFrameOptions = {
        type: "graphicFrame",
        id: cnvPr.id,
        nonVisualProperties: cnvPr.nvp,
        transform: parse(transform2DDesc, xfrm, ctx),
        graphic: stringifyElement(graphic),
      };
      const cNvGraphicFramePr = findChild(nvGraphicFramePr, "cdr:cNvGraphicFramePr");
      if (cNvGraphicFramePr) {
        const locks = findChild(cNvGraphicFramePr, "a:graphicFrameLocks");
        if (locks) {
          const flags: Record<string, boolean> = {};
          for (const key of [
            "noGrp",
            "noDrilldown",
            "noSelect",
            "noChangeAspect",
            "noMove",
            "noResize",
          ]) {
            const raw = locks.attributes?.[key];
            if (raw !== undefined) flags[key] = parseOnOff(raw) ?? false;
          }
          if (Object.keys(flags).length > 0)
            result.graphicFrameLocks = flags as GraphicFrameLockingOptions;
        } else {
          result.graphicFrameLocks = {};
        }
      }
      Object.assign(result, readCommonAttrs(el));
      return result;
    }
    case "cdr:grpSp": {
      const nvGrpSpPr = findChild(el, "cdr:nvGrpSpPr");
      const grpSpPr = findChild(el, "cdr:grpSpPr");
      if (!nvGrpSpPr || !grpSpPr) return undefined;
      const cnvPr = readCnvPr(findChild(nvGrpSpPr, "cdr:cNvPr"));
      if (!cnvPr) return undefined;
      const children: UserShapeObjectOptions[] = [];
      for (const child of el.elements ?? []) {
        const parsed = readObject(child, ctx);
        if (parsed) children.push(parsed);
      }
      return {
        type: "group",
        id: cnvPr.id,
        nonVisualProperties: cnvPr.nvp,
        groupShapeProperties: parse(groupShapePropertiesDesc, grpSpPr, ctx),
        children,
      };
    }
    default:
      return undefined;
  }
}

// ── Descriptor ──

export const userShapesDesc: CustomDescriptor<UserShapesOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const anchors = opts.anchors
      .map((anchor) => {
        const objectXml = stringifyObject(anchor.object);
        if (isRelative(anchor)) {
          return (
            `<cdr:relSizeAnchor>` +
            `<cdr:from>${stringifyMarker(anchor.from)}</cdr:from>` +
            `<cdr:to>${stringifyMarker(anchor.to)}</cdr:to>` +
            objectXml +
            `</cdr:relSizeAnchor>`
          );
        }
        return (
          `<cdr:absSizeAnchor>` +
          `<cdr:from>${stringifyMarker(anchor.from)}</cdr:from>` +
          `<cdr:ext cx="${anchor.extent.width}" cy="${anchor.extent.height}"/>` +
          objectXml +
          `</cdr:absSizeAnchor>`
        );
      })
      .join("");

    return (
      `<cdr:userShapes xmlns:cdr="${CDR_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">` +
      anchors +
      `</cdr:userShapes>`
    );
  },

  parse(el, ctx) {
    const anchors: (RelativeSizeAnchorOptions | AbsoluteSizeAnchorOptions)[] = [];
    for (const anchorEl of el.elements ?? []) {
      if (anchorEl.name === "cdr:relSizeAnchor") {
        const from = readMarker(findChild(anchorEl, "cdr:from"));
        const to = readMarker(findChild(anchorEl, "cdr:to"));
        const object = readAnchorObject(anchorEl, ctx);
        if (from && to && object) anchors.push({ from, to, object });
      } else if (anchorEl.name === "cdr:absSizeAnchor") {
        const from = readMarker(findChild(anchorEl, "cdr:from"));
        const ext = findChild(anchorEl, "cdr:ext");
        const object = readAnchorObject(anchorEl, ctx);
        if (from && ext && object) {
          anchors.push({
            from,
            extent: {
              width: Number(ext.attributes?.["cx"] ?? 0),
              height: Number(ext.attributes?.["cy"] ?? 0),
            },
            object,
          });
        }
      }
    }
    return { anchors } as UserShapesOptions;
  },
};

/** Read the EG_ObjectChoices payload of an anchor element. */
function readAnchorObject(
  anchorEl: XmlElement,
  ctx: ReadContext,
): UserShapeObjectOptions | undefined {
  for (const child of anchorEl.elements ?? []) {
    const object = readObject(child, ctx);
    if (object) return object;
  }
  return undefined;
}
