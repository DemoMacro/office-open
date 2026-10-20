/**
 * OLE Object descriptor for PPTX.
 *
 * Produces a p:graphicFrame with p:oleObj.
 *
 * @module
 */

import {
  convertToEmu,
  createSourceRectangle,
  fillDesc,
  outlineDesc,
  toUint8Array,
} from "@office-open/core";
import { parse, stringify } from "@office-open/core/descriptor";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { findFillChild } from "@office-open/core/drawing";
import { attr, attrBool, attrNum, escapeXml, findChild, findFirst } from "@office-open/xml";

import type { PptxWriteContext } from "../../context";
import type { OleIconImageOptions, OleOptions } from "../ole-frame";
import {
  readGraphicFrameLocking,
  readGraphicFrameHyperlink,
  stringifyGraphicFrameCnvPr,
  readNvPrPlaceholder,
  stringifyCnvGraphicFramePr,
  stringifyNvPr,
} from "./graphic-frame";
import { readCnvPr, readPositionFromXfrm } from "./shape";
import { nextSlideDrawingId } from "./slide-drawing-ids";

// ── ID counter ──

let _nextOleId = 2048;

// ── OLE descriptor ──

export const oleDesc: CustomDescriptor<OleOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const id = opts.id ?? nextSlideDrawingId(ctx) ?? _nextOleId++;
    const name = opts.name ?? `Object ${id}`;
    (ctx as PptxWriteContext).registerShapeId(name, id);

    const x = convertToEmu(opts.x ?? 0);
    const y = convertToEmu(opts.y ?? 0);
    const w = convertToEmu(opts.width ?? "100px");
    const h = convertToEmu(opts.height ?? "100px");

    const parts: string[] = [];

    // p:nvGraphicFramePr
    parts.push(
      `<p:nvGraphicFramePr>${stringifyGraphicFrameCnvPr(id, opts, name, ctx)}` +
        `${stringifyCnvGraphicFramePr(opts.locking)}` +
        `${stringifyNvPr(opts)}</p:nvGraphicFramePr>`,
    );

    // p:xfrm
    parts.push(`<p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></p:xfrm>`);

    // a:graphic > a:graphicData > p:oleObj
    const oleAttrs: string[] = [];
    if (opts.objectName !== false) {
      oleAttrs.push(`name="${escapeXml(opts.objectName ?? opts.name ?? "OLE Object")}"`);
    }
    if (opts.shapeId !== undefined) oleAttrs.push(`spid="${escapeXml(opts.shapeId)}"`);
    if (opts.showAsIcon) oleAttrs.push(`showAsIcon="1"`);
    if (opts.imageWidth !== undefined) oleAttrs.push(`imgW="${opts.imageWidth}"`);
    if (opts.imageHeight !== undefined) oleAttrs.push(`imgH="${opts.imageHeight}"`);
    if (opts.progId !== undefined) oleAttrs.push(`progId="${escapeXml(opts.progId)}"`);

    // Embedded OLE: register the binary as ppt/embeddings/oleObjectN.bin and
    // emit a {ole:…} placeholder — the compiler rewrites it to a real r:id and
    // adds the oleObject relationship. Linked OLE emits an {ole-link:…}
    // placeholder resolved to an External oleObject relationship the same way.
    // followColorScheme is CT_OleObjectEmbed's attribute, not oleObj's.
    const oleChildren: string[] = [];
    if (opts.embed) {
      const pptxCtx = ctx as PptxWriteContext;
      const ref = pptxCtx.addOle(
        toUint8Array(opts.embed.data) as Uint8Array,
        opts.progId,
        opts.embed.fileName,
        opts.embed.relationshipType,
      );
      oleAttrs.push(`r:id="${ref}"`);
      const fcs = opts.embed.followColorScheme
        ? ` followColorScheme="${opts.embed.followColorScheme}"`
        : "";
      oleChildren.push(`<p:embed${fcs}/>`);
    } else if (opts.link) {
      const pptxCtx = ctx as PptxWriteContext;
      const linkKey = pptxCtx.addOleLink(opts.link.url);
      oleAttrs.push(`r:id="{ole-link:${linkKey}}"`);
      const linkAttrs = opts.link.autoUpdate ? ' updateAutomatic="1"' : "";
      oleChildren.push(`<p:link${linkAttrs}/>`);
    } else {
      // p:oleObj's content model is a required choice between p:embed and
      // p:link — emitting neither produces schema-invalid XML.
      throw new Error("OleOptions requires either embed or link");
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
        opts.iconImage.fileName,
      );
      const stretch =
        opts.iconImage.stretchPresent === false
          ? ""
          : opts.iconImage.fillRectangle === false
            ? "<a:stretch/>"
            : `<a:stretch>${createSourceRectangle(
                opts.iconImage.fillRectangle ?? {},
                "a:fillRect",
              )}</a:stretch>`;
      const iconX = convertToEmu(opts.iconImage.x ?? x);
      const iconY = convertToEmu(opts.iconImage.y ?? y);
      const iconWidth = convertToEmu(opts.iconImage.width ?? w);
      const iconHeight = convertToEmu(opts.iconImage.height ?? h);
      const fillXml =
        opts.iconImage.fill === undefined
          ? ""
          : (stringify(fillDesc, opts.iconImage.fill, ctx) ?? "");
      const outlineXml =
        opts.iconImage.outline === undefined
          ? ""
          : (stringify(outlineDesc, opts.iconImage.outline, ctx) ?? "");
      const hasIconTransform =
        opts.iconImage.shapePropertiesPresent !== false &&
        opts.iconImage.transformPresent !== false;
      const iconGeometry =
        opts.iconImage.geometry === null
          ? ""
          : `<a:prstGeom prst="${opts.iconImage.geometry ?? "rect"}"><a:avLst/></a:prstGeom>`;
      const iconProperties =
        opts.iconImage.shapePropertiesPresent === false
          ? ""
          : `<p:spPr>${
              hasIconTransform
                ? `<a:xfrm><a:off x="${iconX}" y="${iconY}"/><a:ext cx="${iconWidth}" cy="${iconHeight}"/></a:xfrm>`
                : ""
            }${iconGeometry}${fillXml}${outlineXml}</p:spPr>`;
      const iconCnvPr = `<p:cNvPr id="${opts.iconImage.id ?? 0}" name="${escapeXml(
        opts.iconImage.name ?? "",
      )}"/>`;
      oleChildren.push(
        `<p:pic><p:nvPicPr>${iconCnvPr}<p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
          `<p:blipFill><a:blip r:embed="${imageRef}"` +
          `${opts.iconImage.compression ? ` cstate="${opts.iconImage.compression}"` : ""}/>` +
          `${stretch}</p:blipFill>${iconProperties}</p:pic>`,
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
    Object.assign(result, readCnvPr(el, "p:nvGraphicFramePr"));
    const locking = readGraphicFrameLocking(findChild(el, "p:nvGraphicFramePr"), ctx);
    const hyperlink = readGraphicFrameHyperlink(findChild(el, "p:nvGraphicFramePr"), ctx);
    readNvPrPlaceholder(findChild(el, "p:nvGraphicFramePr") ?? el, result);
    if (hyperlink) result.hyperlink = hyperlink;
    if (locking !== undefined) result.locking = locking;

    // x, y, width, height from p:xfrm (in EMU)
    const xfrm = findChild(el, "p:xfrm");
    if (xfrm) Object.assign(result, readPositionFromXfrm(xfrm));

    // Navigate to a:graphic/a:graphicData/p:oleObj
    const graphic = findChild(el, "a:graphic");
    const graphicData = graphic ? findChild(graphic, "a:graphicData") : undefined;
    // Model the canonical markup-compatibility branch, not every branch. The
    // fallback p:pic is presentation-specific recovery markup, while the
    // canonical branch intentionally omits it; reading both would turn it into
    // output-only content on regeneration.
    const alternateContent = graphicData
      ? findChild(graphicData, "mc:AlternateContent")
      : undefined;
    const canonicalBranch = alternateContent ? findChild(alternateContent, "mc:Choice") : undefined;
    const canonicalOleObj = canonicalBranch
      ? findChild(canonicalBranch, "p:oleObj")
      : graphicData
        ? findChild(graphicData, "p:oleObj")
        : undefined;
    for (const oleObj of canonicalOleObj ? [canonicalOleObj] : []) {
      const progId = attr(oleObj, "progId");
      if (progId !== undefined && result.progId === undefined) result.progId = progId;
      const objectName = attr(oleObj, "name");
      if (objectName !== undefined && result.objectName === undefined)
        result.objectName = objectName;
      const shapeId = attr(oleObj, "spid");
      if (shapeId !== undefined && result.shapeId === undefined) result.shapeId = shapeId;
      if (attrBool(oleObj, "showAsIcon")) result.showAsIcon = true;
      const imgW = attrNum(oleObj, "imgW");
      if (imgW !== undefined && result.imageWidth === undefined) result.imageWidth = imgW;
      const imgH = attrNum(oleObj, "imgH");
      if (imgH !== undefined && result.imageHeight === undefined) result.imageHeight = imgH;
      result.objectName = objectName ?? false;

      // embed/link — embedded OLE reads the binary back through the
      // relationship so generate re-registers it in a fresh package.
      const embedEl = findChild(oleObj, "p:embed");
      if (embedEl) {
        const rId = attr(oleObj, "r:id");
        const mediaPath = rId ? ctx.resolveRelationship(rId) : undefined;
        const raw = mediaPath ? ctx.getRaw(mediaPath) : undefined;
        const followCS = attr(embedEl, "followColorScheme") as
          | "none"
          | "full"
          | "textAndBackground"
          | undefined;
        const sourceFileName = mediaPath?.split("/").pop();
        if (raw) {
          const relationshipType = rId ? ctx.resolveEmbeddingType?.(rId) : undefined;
          result.embed = {
            data: raw,
            ...(sourceFileName !== undefined ? { fileName: sourceFileName } : {}),
            ...(relationshipType !== undefined ? { relationshipType } : {}),
            ...(followCS !== undefined ? { followColorScheme: followCS } : {}),
          };
        }
      } else {
        const linkEl = findChild(oleObj, "p:link");
        if (linkEl) {
          const rId = attr(oleObj, "r:id");
          // The linked source resolves through the External oleObject
          // relationship — the URL is the durable fact; the source rId itself
          // does not survive renumbering.
          const url = rId ? ctx.resolveRelationship(rId) : undefined;
          const autoUpdate = attrBool(linkEl, "updateAutomatic") === true;
          if (url && result.link === undefined)
            result.link = { url, ...(autoUpdate ? { autoUpdate: true } : {}) };
        }
      }

      // iconImage: read the picture bytes back through the blip relationship
      const pic = findChild(oleObj, "p:pic");
      if (pic) {
        const iconCnvPr = findChild(pic, "p:nvPicPr")
          ? findChild(findChild(pic, "p:nvPicPr")!, "p:cNvPr")
          : undefined;
        const iconImageId = iconCnvPr ? attrNum(iconCnvPr, "id") : undefined;
        const iconImageName = iconCnvPr ? attr(iconCnvPr, "name") : undefined;
        const blip = findFirst(pic, "a:blip");
        const blipRId = blip ? attr(blip, "r:embed") : undefined;
        const imagePath = blipRId ? ctx.resolveRelationship(blipRId) : undefined;
        const raw = imagePath ? ctx.getRaw(imagePath) : undefined;
        const type = imagePath?.split(".").pop();
        const compression = blip ? attr(blip, "cstate") : undefined;
        const iconBlipFill = findChild(pic, "p:blipFill");
        const hasStretch =
          iconBlipFill !== undefined && findChild(iconBlipFill, "a:stretch") !== undefined;
        const iconStretch =
          iconBlipFill === undefined ? undefined : findChild(iconBlipFill, "a:stretch");
        const fillRect =
          iconStretch === undefined ? undefined : findChild(iconStretch, "a:fillRect");
        const iconSpPr = findChild(pic, "p:spPr");
        const iconFillChild = iconSpPr ? findFillChild(iconSpPr) : undefined;
        const iconFill = iconFillChild ? parse(fillDesc, iconFillChild, ctx) : undefined;
        const iconLine = iconSpPr ? findChild(iconSpPr, "a:ln") : undefined;
        const iconOutline = iconLine ? parse(outlineDesc, iconLine, ctx) : undefined;
        const iconXfrm = iconSpPr ? findChild(iconSpPr, "a:xfrm") : undefined;
        const iconOff = iconXfrm ? findChild(iconXfrm, "a:off") : undefined;
        const iconExt = iconXfrm ? findChild(iconXfrm, "a:ext") : undefined;
        const iconGeometry = iconSpPr ? findChild(iconSpPr, "a:prstGeom") : undefined;
        const iconGeometryType = iconGeometry ? attr(iconGeometry, "prst") : undefined;
        if (raw && type && result.iconImage === undefined) {
          result.iconImage = {
            data: raw,
            type,
            ...(iconImageId !== undefined ? { id: iconImageId } : {}),
            ...(iconImageName !== undefined ? { name: iconImageName } : {}),
            ...(imagePath !== undefined ? { fileName: imagePath.split("/").pop() } : {}),
            ...(hasStretch ? {} : { stretchPresent: false }),
            ...(hasStretch && fillRect === undefined ? { fillRectangle: false } : {}),
            ...(iconFill !== undefined ? { fill: iconFill } : {}),
            ...(iconOutline !== undefined ? { outline: iconOutline } : {}),
            ...(iconXfrm === undefined ? { transformPresent: false } : {}),
            ...(iconSpPr === undefined ? { shapePropertiesPresent: false } : {}),
            ...(iconOff?.attributes?.x !== undefined ? { x: attrNum(iconOff, "x") } : {}),
            ...(iconOff?.attributes?.y !== undefined ? { y: attrNum(iconOff, "y") } : {}),
            ...(iconExt?.attributes?.cx !== undefined ? { width: attrNum(iconExt, "cx") } : {}),
            ...(iconExt?.attributes?.cy !== undefined ? { height: attrNum(iconExt, "cy") } : {}),
            ...(iconGeometry === undefined
              ? { geometry: null }
              : iconGeometryType !== undefined
                ? { geometry: iconGeometryType as OleIconImageOptions["geometry"] }
                : {}),
            ...(compression !== undefined
              ? { compression: compression as OleIconImageOptions["compression"] }
              : {}),
          };
        }
      }
    }

    return result as OleOptions;
  },
};
