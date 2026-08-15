/**
 * OLE Object descriptor for PPTX.
 *
 * Produces a p:graphicFrame with p:oleObj.
 *
 * @module
 */

import { convertToEmu, toUint8Array } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import {
  stringifyNonVisualDrawingProperties,
  parseNonVisualDrawingProperties,
} from "@office-open/core/drawing";
import { attr, attrBool, attrNum, escapeXml, findChild, type Element } from "@office-open/xml";

import type { PptxWriteContext } from "../../context";
import type { OleOptions } from "../ole-frame";

// ── ID counter ──

let _nextOleId = 2048;

// ── OLE descriptor ──

export const oleDesc: CustomDescriptor<OleOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const id = opts.id ?? _nextOleId++;
    const name = opts.name ?? `Object ${id}`;

    const x = convertToEmu(opts.x ?? 0);
    const y = convertToEmu(opts.y ?? 0);
    const w = convertToEmu(opts.width ?? "100px");
    const h = convertToEmu(opts.height ?? "100px");

    const parts: string[] = [];

    // p:nvGraphicFramePr
    parts.push(
      `<p:nvGraphicFramePr>${stringifyNonVisualDrawingProperties("p:cNvPr", id, opts, name)}` +
        `<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>` +
        `<p:nvPr/></p:nvGraphicFramePr>`,
    );

    // p:xfrm
    parts.push(`<p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></p:xfrm>`);

    // a:graphic > a:graphicData > p:oleObj
    const oleAttrs: string[] = [];
    oleAttrs.push(`name="${escapeXml(opts.name ?? "OLE Object")}"`);
    if (opts.shapeId) oleAttrs.push(`spid="${opts.shapeId}"`);
    if (opts.showAsIcon) oleAttrs.push(`showAsIcon="1"`);
    if (opts.imgW !== undefined) oleAttrs.push(`imgW="${opts.imgW}"`);
    if (opts.imgH !== undefined) oleAttrs.push(`imgH="${opts.imgH}"`);
    if (opts.progId) oleAttrs.push(`progId="${opts.progId}"`);

    // Embedded OLE: register the binary as ppt/embeddings/oleObjectN.bin and
    // emit a {ole:…} placeholder — the compiler rewrites it to a real r:id and
    // adds the oleObject relationship. Linked OLE keeps its external rId as-is.
    // followColorScheme is CT_OleObjectEmbed's attribute, not oleObj's.
    const oleChildren: string[] = [];
    if (opts.embed) {
      const pptxCtx = ctx as PptxWriteContext;
      const ref = pptxCtx.addOle(toUint8Array(opts.embed.data) as Uint8Array, opts.progId);
      oleAttrs.push(`r:id="${ref}"`);
      const fcs = opts.followColorScheme ? ` followColorScheme="${opts.followColorScheme}"` : "";
      oleChildren.push(`<p:embed${fcs}/>`);
    } else if (opts.link) {
      oleAttrs.push(`r:id="${opts.link.rId}"`);
      const linkAttrs = opts.link.autoUpdate ? ' updateAutomatic="1"' : "";
      oleChildren.push(`<p:link${linkAttrs}/>`);
    }

    // Icon/preview picture — MS Office refuses to open the presentation when
    // an oleObj carries no picture, so it is emitted whenever iconImage is
    // supplied. Registered as media; the {fileName} placeholder is rewritten
    // by the compiler's image pass.
    if (opts.iconImage) {
      const pptxCtx = ctx as PptxWriteContext;
      const imageRef = pptxCtx.addMedia(
        toUint8Array(opts.iconImage.data) as Uint8Array,
        opts.iconImage.type,
      );
      oleChildren.push(
        `<p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
          `<p:blipFill><a:blip r:embed="${imageRef}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
          `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>` +
          `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`,
      );
    }

    parts.push(
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">` +
        `<p:oleObj ${oleAttrs.join(" ")}>${oleChildren.join("")}</p:oleObj>` +
        `</a:graphicData></a:graphic>`,
    );

    return `<p:graphicFrame>${parts.join("")}</p:graphicFrame>`;
  },

  parse(el, ctx) {
    const result: Partial<OleOptions> = {};

    // id, name from p:nvGraphicFramePr/p:cNvPr
    const nvGrFrm = findChild(el, "p:nvGraphicFramePr");
    if (nvGrFrm) {
      const cnvPr = findChild(nvGrFrm, "p:cNvPr");
      if (cnvPr) {
        Object.assign(result, parseNonVisualDrawingProperties(cnvPr));
        const id = attrNum(cnvPr, "id");
        if (id !== undefined) result.id = id;
      }
    }

    // x, y, width, height from p:xfrm (in EMU)
    const xfrm = findChild(el, "p:xfrm");
    if (xfrm) {
      const off = findChild(xfrm, "a:off");
      if (off) {
        const x = attrNum(off, "x");
        if (x !== undefined) result.x = x;
        const y = attrNum(off, "y");
        if (y !== undefined) result.y = y;
      }
      const ext = findChild(xfrm, "a:ext");
      if (ext) {
        const cx = attrNum(ext, "cx");
        if (cx !== undefined) result.width = cx;
        const cy = attrNum(ext, "cy");
        if (cy !== undefined) result.height = cy;
      }
    }

    // Navigate to a:graphic/a:graphicData/p:oleObj
    const graphic = findChild(el, "a:graphic");
    const graphicData = graphic ? findChild(graphic, "a:graphicData") : undefined;
    const oleObj = graphicData ? findChild(graphicData, "p:oleObj") : undefined;
    if (oleObj) {
      const progId = attr(oleObj, "progId");
      if (progId !== undefined) result.progId = progId;
      const shapeId = attr(oleObj, "spid");
      if (shapeId !== undefined) result.shapeId = shapeId;
      if (attrBool(oleObj, "showAsIcon")) result.showAsIcon = true;
      const imgW = attrNum(oleObj, "imgW");
      if (imgW !== undefined) result.imgW = imgW;
      const imgH = attrNum(oleObj, "imgH");
      if (imgH !== undefined) result.imgH = imgH;

      // embed/link — embedded OLE reads the binary back through the
      // relationship so generate re-registers it in a fresh package.
      const embedEl = findChild(oleObj, "p:embed");
      if (embedEl) {
        const rId = attr(oleObj, "r:id");
        const mediaPath = rId ? ctx.resolveRelationship(rId) : undefined;
        const raw = mediaPath ? ctx.getRaw(mediaPath) : undefined;
        if (raw) result.embed = { data: raw };
        const followCS = attr(embedEl, "followColorScheme");
        if (followCS !== undefined)
          result.followColorScheme = followCS as "none" | "full" | "textAndBackground";
      } else {
        const linkEl = findChild(oleObj, "p:link");
        if (linkEl) {
          const rId = attr(oleObj, "r:id");
          const autoUpdate = attrBool(linkEl, "updateAutomatic") === true;
          result.link = { rId: rId ?? "", ...(autoUpdate ? { autoUpdate: true } : {}) };
        }
      }

      // iconImage: read the picture bytes back through the blip relationship
      const pic = findChild(oleObj, "p:pic");
      if (pic) {
        const blip = findChildDeep(pic, "a:blip");
        const blipRId = blip ? attr(blip, "r:embed") : undefined;
        const imagePath = blipRId ? ctx.resolveRelationship(blipRId) : undefined;
        const raw = imagePath ? ctx.getRaw(imagePath) : undefined;
        const type = imagePath?.split(".").pop();
        if (raw && type) result.iconImage = { data: raw, type };
      }
    }

    return result as OleOptions;
  },
};

/** Find a descendant element by name at any depth. */
function findChildDeep(el: Element, name: string): Element | undefined {
  for (const child of el.elements ?? []) {
    if (child.name === name) return child;
    const found = findChildDeep(child, name);
    if (found) return found;
  }
  return undefined;
}
