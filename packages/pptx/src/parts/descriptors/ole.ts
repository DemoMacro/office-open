/**
 * OLE Object descriptor for PPTX.
 *
 * Produces a p:graphicFrame with p:oleObj.
 *
 * @module
 */

import { convertEmuToPixels, convertPixelsToEmu } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrBool, attrNum, escapeXml, findChild } from "@office-open/xml";

// ── Types ──

export interface OleDescriptorOptions {
  id?: number;
  name?: string;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  progId?: string;
  spid?: string;
  showAsIcon?: boolean;
  imgW?: number;
  imgH?: number;
  embed?: { rId: string; lastEdited?: string };
  link?: { rId: string; updateLastEdited?: string; autoUpdate?: boolean };
  imgRId?: string;
  followColorScheme?: "none" | "full" | "textAndBackground";
}

// ── ID counter ──

let _nextOleId = 2048;

// ── OLE descriptor ──

export const oleDesc: CustomDescriptor<OleDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const id = opts.id ?? _nextOleId++;
    const name = opts.name ?? `Object ${id}`;

    const x = convertPixelsToEmu(opts.x ?? 0);
    const y = convertPixelsToEmu(opts.y ?? 0);
    const w = convertPixelsToEmu(opts.width ?? 100);
    const h = convertPixelsToEmu(opts.height ?? 100);

    const parts: string[] = [];

    // p:nvGraphicFramePr
    parts.push(
      `<p:nvGraphicFramePr><p:cNvPr id="${id}" name="${escapeXml(name)}"/>` +
        `<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>` +
        `<p:nvPr/></p:nvGraphicFramePr>`,
    );

    // p:xfrm
    parts.push(`<p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></p:xfrm>`);

    // a:graphic > a:graphicData > p:oleObj
    const oleAttrs: string[] = [];
    oleAttrs.push(`name="${escapeXml(opts.name ?? "OLE Object")}"`);
    if (opts.spid) oleAttrs.push(`spid="${opts.spid}"`);
    if (opts.showAsIcon) oleAttrs.push(`showAsIcon="1"`);
    if (opts.imgW !== undefined) oleAttrs.push(`imgW="${opts.imgW}"`);
    if (opts.imgH !== undefined) oleAttrs.push(`imgH="${opts.imgH}"`);
    if (opts.progId) oleAttrs.push(`progId="${opts.progId}"`);
    if (opts.followColorScheme) oleAttrs.push(`followColorScheme="${opts.followColorScheme}"`);
    const rId = opts.embed?.rId ?? opts.link?.rId;
    if (rId) oleAttrs.push(`r:id="${rId}"`);

    const oleChildren: string[] = [];
    if (opts.embed) {
      oleChildren.push("<p:embed/>");
    } else if (opts.link) {
      const linkAttrs = opts.link.autoUpdate ? ' updateAutomatic="1"' : "";
      oleChildren.push(`<p:link${linkAttrs}/>`);
    }

    if (opts.imgRId) {
      oleChildren.push(
        `<p:pic><p:nvPicPr><p:cNvPr id="2" name="${escapeXml(name)}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr></p:pic>`,
      );
    }

    parts.push(
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">` +
        `<p:oleObj ${oleAttrs.join(" ")}>${oleChildren.join("")}</p:oleObj>` +
        `</a:graphicData></a:graphic>`,
    );

    return `<p:graphicFrame>${parts.join("")}</p:graphicFrame>`;
  },

  parse(el, _ctx) {
    const result: Partial<OleDescriptorOptions> = {};

    // id, name from p:nvGraphicFramePr/p:cNvPr
    const nvGrFrm = findChild(el, "p:nvGraphicFramePr");
    if (nvGrFrm) {
      const cnvPr = findChild(nvGrFrm, "p:cNvPr");
      if (cnvPr) {
        const id = attrNum(cnvPr, "id");
        if (id !== undefined) result.id = id;
        const name = attr(cnvPr, "name");
        if (name !== undefined) result.name = name;
      }
    }

    // x, y, width, height from p:xfrm (convert EMU to pixels)
    const xfrm = findChild(el, "p:xfrm");
    if (xfrm) {
      const off = findChild(xfrm, "a:off");
      if (off) {
        const x = attrNum(off, "x");
        if (x !== undefined) result.x = convertEmuToPixels(x);
        const y = attrNum(off, "y");
        if (y !== undefined) result.y = convertEmuToPixels(y);
      }
      const ext = findChild(xfrm, "a:ext");
      if (ext) {
        const cx = attrNum(ext, "cx");
        if (cx !== undefined) result.width = convertEmuToPixels(cx);
        const cy = attrNum(ext, "cy");
        if (cy !== undefined) result.height = convertEmuToPixels(cy);
      }
    }

    // Navigate to a:graphic/a:graphicData/p:oleObj
    const graphic = findChild(el, "a:graphic");
    const graphicData = graphic ? findChild(graphic, "a:graphicData") : undefined;
    const oleObj = graphicData ? findChild(graphicData, "p:oleObj") : undefined;
    if (oleObj) {
      const progId = attr(oleObj, "progId");
      if (progId !== undefined) result.progId = progId;
      const spid = attr(oleObj, "spid");
      if (spid !== undefined) result.spid = spid;
      if (attrBool(oleObj, "showAsIcon")) result.showAsIcon = true;
      const imgW = attrNum(oleObj, "imgW");
      if (imgW !== undefined) result.imgW = imgW;
      const imgH = attrNum(oleObj, "imgH");
      if (imgH !== undefined) result.imgH = imgH;
      const followCS = attr(oleObj, "followColorScheme");
      if (followCS !== undefined)
        result.followColorScheme = followCS as "none" | "full" | "textAndBackground";

      // embed/link
      const embedEl = findChild(oleObj, "p:embed");
      if (embedEl) {
        const rId = attr(oleObj, "r:id");
        if (rId !== undefined) result.embed = { rId };
      } else {
        const linkEl = findChild(oleObj, "p:link");
        if (linkEl) {
          const rId = attr(oleObj, "r:id");
          const autoUpdate = attrBool(linkEl, "updateAutomatic") === true;
          result.link = { rId: rId ?? "", ...(autoUpdate ? { autoUpdate: true } : {}) };
        }
      }

      // imgRId: search for r:id in p:pic subtree
      const pic = findChild(oleObj, "p:pic");
      if (pic) {
        const picRId = findDeepAttr(pic, "r:id");
        if (picRId !== undefined) result.imgRId = picRId;
      }
    }

    return result as OleDescriptorOptions;
  },
};

/** Search for an attribute value in the element or any descendant. */
function findDeepAttr(
  el: {
    attributes?: Record<string, string | number | undefined>;
    elements?: Array<{
      attributes?: Record<string, string | number | undefined>;
      elements?: Array<unknown>;
    }>;
  },
  name: string,
): string | undefined {
  const v = el.attributes?.[name];
  if (v !== undefined) return String(v);
  for (const child of el.elements ?? []) {
    const found = findDeepAttr(child as Parameters<typeof findDeepAttr>[0], name);
    if (found !== undefined) return found;
  }
  return undefined;
}
