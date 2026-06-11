/**
 * OLE Object descriptor for PPTX.
 *
 * Produces a p:graphicFrame with p:oleObj.
 *
 * @module
 */

import { convertPixelsToEmu } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { escapeXml } from "@office-open/xml";

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

  parse(_el, _ctx) {
    return {};
  },
};
