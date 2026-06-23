/**
 * Object element for WordprocessingML documents — w:object.
 *
 * Embeds an OLE object (e.g. an Excel sheet) in a run via a VML preview shape and
 * exactly one of objectEmbed / objectLink / control / movie. The OLE binary is
 * registered as word/embeddings/oleObjectN.bin (EmbeddingCollection); the optional
 * preview icon as word/media/imageN.<type> (Media). Relationship ids are emitted
 * as `{fileName}` placeholders and rewritten by the compiler's media bridge.
 *
 * Reference: OOXML transitional, wml.xsd, CT_Object / CT_ObjectEmbed / CT_ObjectLink
 *
 * @module
 */
import { toUint8Array } from "@office-open/core";
import type { UniversalMeasure } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrNum, findChild, type Element } from "@office-open/xml";
import type { EmbeddingData } from "@shared/embeddings/embeddings";
import type { MediaData } from "@shared/media/data";

import type { BodyContext } from "../../context";
import { createImageData } from "../paragraph/run/image-run";

// ── Options ──

export interface ObjectEmbedOptions {
  /** OLE container binary — registered as word/embeddings/oleObjectN.bin. */
  data: Uint8Array | string;
  /** OLE program id (e.g. "Excel.Sheet.12"). */
  progId?: string;
  /** Draw aspect — how the object displays. */
  drawAspect?: "content" | "icon";
  /** Shape id (w:objectEmbed/@shapeId). */
  shapeId?: string;
  /** Field codes (w:objectEmbed/@fieldCodes). */
  fieldCodes?: string;
}

export interface ObjectLinkOptions extends ObjectEmbedOptions {
  /** Update mode (required for links). */
  updateMode: "always" | "onCall";
  /** Whether the field is locked. */
  lockedField?: boolean;
}

export interface ObjectControlOptions {
  /** Control name (w:control/@name). */
  name?: string;
  /** Shape id (w:control/@shapeid). */
  shapeid?: string;
  /** Relationship id to the ActiveX part (external — not auto-registered). */
  rId: string;
}

export interface ObjectIconImageOptions {
  /** Preview image bytes (binary or base64 data URL). */
  data: Uint8Array | string;
  /** Image type / extension (e.g. "png", "emf"). */
  type: string;
  /** Title for v:imagedata/@o:title. */
  title?: string;
}

export interface ObjectElementOptions {
  /** Original width in twips (w:object/@w:dxaOrig). */
  dxaOrig?: number;
  /** Original height in twips (w:object/@w:dyaOrig). */
  dyaOrig?: number;
  /** VML shape id (v:shape/@id). Defaults to a generated id. */
  shapeId?: string;
  /** Display width (px or universal measure) for v:shape style + icon size. */
  width?: number | UniversalMeasure;
  /** Display height (px or universal measure). */
  height?: number | UniversalMeasure;
  /** Preview icon image (v:imagedata). */
  iconImage?: ObjectIconImageOptions;
  /** Embedded OLE object (w:objectEmbed). */
  embed?: ObjectEmbedOptions;
  /** Linked OLE object (w:objectLink). */
  link?: ObjectLinkOptions;
  /** ActiveX control reference (w:control). */
  control?: ObjectControlOptions;
  /** Movie relationship id — CT_Rel (w:movie/@r:id). External. */
  movie?: string;
}

// ── Descriptor ──

let objectShapeCounter = 1025;

export const objectDesc: CustomDescriptor<ObjectElementOptions, BodyContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    const inner: string[] = [];

    // VML preview shape (v:shape + optional v:imagedata)
    const shapeId = opts.shapeId ?? `_x0000_i${objectShapeCounter++}`;
    const widthVal = opts.width ?? 100;
    const heightVal = opts.height ?? 100;
    const styleWidth = typeof widthVal === "number" ? `${widthVal}px` : widthVal;
    const styleHeight = typeof heightVal === "number" ? `${heightVal}px` : heightVal;

    const shapeChildren: string[] = [];
    if (opts.iconImage) {
      const iconFileName = ctx.file.media.nextMediaName(opts.iconImage.type);
      const iconMediaData = {
        type: opts.iconImage.type,
        ...createImageData(
          toUint8Array(opts.iconImage.data) as Uint8Array,
          { width: widthVal, height: heightVal },
          iconFileName,
        ),
      } as MediaData;
      ctx.file.media.addImage(iconFileName, iconMediaData);
      const titleAttr = opts.iconImage.title ? ` o:title="${opts.iconImage.title}"` : "";
      shapeChildren.push(`<v:imagedata r:id="{${iconFileName}}"${titleAttr}/>`);
    }
    inner.push(
      `<v:shape id="${shapeId}" type="#_x0000_t75" style="width:${styleWidth};height:${styleHeight}">${shapeChildren.join("")}</v:shape>`,
    );

    // Choice: objectEmbed | objectLink | control | movie
    if (opts.embed) {
      const fileName = registerEmbedding(opts.embed, ctx);
      inner.push(`<w:objectEmbed r:id="{${fileName}}"${embedAttrs(opts.embed)}/>`);
    } else if (opts.link) {
      const fileName = registerEmbedding(opts.link, ctx);
      const locked = opts.link.lockedField ? ` w:lockedField="true"` : "";
      inner.push(
        `<w:objectLink r:id="{${fileName}}"${embedAttrs(opts.link)} w:updateMode="${opts.link.updateMode}"${locked}/>`,
      );
    } else if (opts.control) {
      const c = opts.control;
      const cAttrs: string[] = [` r:id="${c.rId}"`];
      if (c.name) cAttrs.push(` w:name="${c.name}"`);
      if (c.shapeid) cAttrs.push(` w:shapeid="${c.shapeid}"`);
      inner.push(`<w:control${cAttrs.join("")}/>`);
    } else if (opts.movie) {
      inner.push(`<w:movie r:id="${opts.movie}"/>`);
    }

    // w:object root attributes
    const objAttrs: string[] = [];
    if (opts.dxaOrig !== undefined) objAttrs.push(` w:dxaOrig="${opts.dxaOrig}"`);
    if (opts.dyaOrig !== undefined) objAttrs.push(` w:dyaOrig="${opts.dyaOrig}"`);

    return `<w:object${objAttrs.join("")}>${inner.join("")}</w:object>`;
  },

  parse(el, _ctx) {
    const result: Partial<ObjectElementOptions> = {};

    const dxaOrig = attrNum(el, "w:dxaOrig");
    if (dxaOrig !== undefined) result.dxaOrig = dxaOrig;
    const dyaOrig = attrNum(el, "w:dyaOrig");
    if (dyaOrig !== undefined) result.dyaOrig = dyaOrig;

    // VML shape — best-effort structural capture (binary media is not re-registered on parse)
    const shape = findChild(el, "v:shape");
    if (shape) {
      const id = attr(shape, "id");
      if (id) result.shapeId = id;
      const style = attr(shape, "style");
      if (style) {
        const w = style.match(/width:([^;]+)/);
        const h = style.match(/height:([^;]+)/);
        if (w) result.width = w[1].trim() as UniversalMeasure;
        if (h) result.height = h[1].trim() as UniversalMeasure;
      }
    }

    // Choice elements
    const embedEl = findChild(el, "w:objectEmbed");
    if (embedEl) result.embed = parseEmbed(embedEl);

    const linkEl = findChild(el, "w:objectLink");
    if (linkEl) {
      const base = parseEmbed(linkEl);
      const updateMode = attr(linkEl, "w:updateMode");
      const lockedField = attr(linkEl, "w:lockedField");
      result.link = {
        ...base,
        ...(updateMode ? { updateMode: updateMode as "always" | "onCall" } : {}),
        ...(lockedField !== undefined
          ? { lockedField: lockedField === "true" || lockedField === "1" }
          : {}),
      } as ObjectLinkOptions;
    }

    const controlEl = findChild(el, "w:control");
    if (controlEl) {
      const rId = attr(controlEl, "r:id") ?? "";
      const name = attr(controlEl, "w:name");
      const shapeid = attr(controlEl, "w:shapeid");
      result.control = { rId, ...(name ? { name } : {}), ...(shapeid ? { shapeid } : {}) };
    }

    const movieEl = findChild(el, "w:movie");
    if (movieEl) {
      const rId = attr(movieEl, "r:id");
      if (rId) result.movie = rId;
    }

    return result as ObjectElementOptions;
  },
};

// ── Helpers ──

/** Register an OLE embedding and return its allocated file name. */
function registerEmbedding(opts: ObjectEmbedOptions, ctx: BodyContext): string {
  const fileName = ctx.file.embeddings.nextEmbeddingName();
  const data: EmbeddingData = {
    fileName,
    data: toUint8Array(opts.data) as Uint8Array,
    ...(opts.progId ? { progId: opts.progId } : {}),
  };
  ctx.file.embeddings.addEmbedding(fileName, data);
  return fileName;
}

/** Build the common objectEmbed/objectLink attribute string (excludes r:id). */
function embedAttrs(opts: ObjectEmbedOptions): string {
  const attrs: string[] = [];
  if (opts.drawAspect) attrs.push(` w:drawAspect="${opts.drawAspect}"`);
  if (opts.progId) attrs.push(` w:progId="${opts.progId}"`);
  if (opts.shapeId) attrs.push(` w:shapeId="${opts.shapeId}"`);
  if (opts.fieldCodes) attrs.push(` w:fieldCodes="${opts.fieldCodes}"`);
  return attrs.join("");
}

/** Parse common objectEmbed/objectLink attributes (excludes r:id — external on parse). */
function parseEmbed(el: Element): ObjectEmbedOptions {
  const opts: Partial<ObjectEmbedOptions> = {};
  const drawAspect = attr(el, "w:drawAspect");
  if (drawAspect === "content" || drawAspect === "icon") opts.drawAspect = drawAspect;
  const progId = attr(el, "w:progId");
  if (progId) opts.progId = progId;
  const shapeId = attr(el, "w:shapeId");
  if (shapeId) opts.shapeId = shapeId;
  const fieldCodes = attr(el, "w:fieldCodes");
  if (fieldCodes) opts.fieldCodes = fieldCodes;
  // data is not recoverable from the relationship on parse; callers re-supply it.
  opts.data = new Uint8Array();
  return opts as ObjectEmbedOptions;
}
