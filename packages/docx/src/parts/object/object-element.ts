/**
 * Object element for WordprocessingML documents — w:object.
 *
 * Embeds an OLE object (e.g. an Excel sheet) in a run via a VML preview shape and
 * exactly one of OLEObject / control / movie. The OLE payload is emitted as
 * `<o:OLEObject>` (the vml-officeDrawing element Word actually writes via
 * CT_Object's lax any slot) — the XSD-declared `w:objectEmbed`/`w:objectLink`
 * spellings parse back but are rejected by Word's reader. The OLE binary is
 * registered as word/embeddings/oleObjectN.bin (EmbeddingCollection); the optional
 * preview icon as word/media/imageN.<type> (Media). Relationship ids are emitted
 * as `{fileName}` placeholders and rewritten by the compiler's media bridge.
 *
 * Reference: OOXML transitional, wml.xsd CT_Object; vml-officeDrawing.xsd CT_OLEObject
 *
 * @module
 */
import { toUint8Array, parseOnOff } from "@office-open/core";
import type { UniversalMeasure } from "@office-open/core";
import { parseVmlStyle } from "@office-open/core";
import { stringifyVmlShape, stringifyVmlShapetype, parseVmlShapetype } from "@office-open/core";
import type { VmlShapetypeOptions } from "@office-open/core";
import { parseVmlImageData, type VmlImageDataOptions } from "@office-open/core";
import type { CustomDescriptor, ReadContext } from "@office-open/core/descriptor";
import { attr, attrNum, escapeXml, findChild, textOf, type Element } from "@office-open/xml";
import type { EmbeddingData } from "@shared/embeddings/embeddings";
import type { MediaData } from "@shared/media/data";

import type { BodyContext } from "../../context";
import { createPictureData } from "../paragraph/run/picture-run";

// ── Options ──

export interface ObjectEmbedOptions {
  /** OLE container binary — registered as word/embeddings/oleObjectN.bin. */
  data: Uint8Array | string;
  /** OLE program id (e.g. "Excel.Sheet.12"). */
  progId?: string;
  /** Draw aspect — how the object displays. */
  drawAspect?: "content" | "icon";
  /** Shape id (o:OLEObject/`@ShapeID`). Defaults to the preview shape's id. */
  shapeId?: string;
  /** OLE object id (o:OLEObject/`@ObjectID`). Defaults to a generated id. */
  objectId?: string;
  /** Field codes (o:OLEObject/o:FieldCodes child). */
  fieldCodes?: string;
}

export interface ObjectLinkOptions extends ObjectEmbedOptions {
  /** Update mode (required for links). */
  updateMode: "always" | "onCall";
  /** Whether the field is locked. */
  lockedField?: boolean;
}

export interface ObjectControlOptions {
  /** Control name (w:control/`@name`). */
  name?: string;
  /** Shape id (w:control/`@shapeid`). */
  shapeid?: string;
  /** Relationship id to the ActiveX part (external — not auto-registered). */
  rId: string;
}

export interface ObjectIconImageOptions {
  /** Preview image bytes (binary or base64 data URL). */
  data: Uint8Array | string;
  /** Image type / extension (e.g. "png", "emf"). */
  type: string;
  /** Title for v:imagedata/`@o:title`. */
  title?: string;
}

export interface ObjectElementOptions {
  /** Original width in twips (w:object/`@w:dxaOrig`). */
  dxaOrig?: number;
  /** Original height in twips (w:object/`@w:dyaOrig`). */
  dyaOrig?: number;
  /** VML shape id (v:shape/`@id`). Defaults to a generated id. */
  shapeId?: string;
  /** Display width (px or universal measure) for v:shape style + icon size. */
  width?: number | UniversalMeasure;
  /** Display height (px or universal measure). */
  height?: number | UniversalMeasure;
  /**
   * v:shapetype preamble — Word writes the _x0000_t75 OLE shapetype (with its
   * formula table) before the preview v:shape; round-tripped verbatim.
   */
  shapetype?: VmlShapetypeOptions;
  /** Preview icon image (v:imagedata). */
  iconImage?: ObjectIconImageOptions;
  /** Embedded OLE object (o:OLEObject Type="Embed"). */
  embed?: ObjectEmbedOptions;
  /** Linked OLE object (o:OLEObject Type="Link"). */
  link?: ObjectLinkOptions;
  /** ActiveX control reference (w:control). */
  control?: ObjectControlOptions;
  /** Movie relationship id — CT_Rel (w:movie/`@r:id`). External. */
  movie?: string;
}

// ── Descriptor ──

let objectShapeCounter = 1025;
let objectOleCounter = 1;

export const objectDesc: CustomDescriptor<ObjectElementOptions, BodyContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    const inner: string[] = [];

    // VML preview shape (v:shape + optional v:imagedata)
    const shapeId = opts.shapeId ?? `_x0000_i${objectShapeCounter++}`;
    const widthVal = opts.width ?? 100;
    const heightVal = opts.height ?? 100;
    const styleWidth =
      typeof widthVal === "number" ? (`${widthVal}px` as UniversalMeasure) : widthVal;
    const styleHeight =
      typeof heightVal === "number" ? (`${heightVal}px` as UniversalMeasure) : heightVal;

    let imagedataOptions: VmlImageDataOptions | undefined;
    if (opts.iconImage) {
      const rawData = toUint8Array(opts.iconImage.data) as Uint8Array;
      const iconType = opts.iconImage.type;
      const { fileName: iconFileName } = ctx.file.media.addMedia(
        rawData,
        iconType,
        (fileName) =>
          ({
            type: iconType,
            ...createPictureData(rawData, { width: widthVal, height: heightVal }, fileName),
          }) as MediaData,
      );
      imagedataOptions = {
        relationshipId: `{${iconFileName}}`,
        officeTitle: opts.iconImage.title,
      };
    }
    if (opts.shapetype) inner.push(stringifyVmlShapetype(opts.shapetype));
    inner.push(
      stringifyVmlShape({
        id: shapeId,
        type: "#_x0000_t75",
        // o:ole marks the shape as an OLE container (Word always writes it here).
        ole: "",
        style: { width: styleWidth, height: styleHeight },
        imagedata: imagedataOptions,
      }),
    );

    // Choice: o:OLEObject (embed/link) | w:control | w:movie
    if (opts.embed || opts.link) {
      const link = opts.link;
      const payload = opts.embed ?? link!;
      const fileName = registerEmbedding(payload, ctx);
      const attrs: string[] = [` Type="${link ? "Link" : "Embed"}"`];
      if (payload.progId) attrs.push(` ProgID="${payload.progId}"`);
      // Word ties the OLE object to its preview shape via ShapeID.
      attrs.push(` ShapeID="${payload.shapeId ?? shapeId}"`);
      if (payload.drawAspect) {
        attrs.push(` DrawAspect="${payload.drawAspect === "icon" ? "Icon" : "Content"}"`);
      }
      attrs.push(` ObjectID="${payload.objectId ?? `_${objectOleCounter++}`}"`);
      attrs.push(` r:id="{${fileName}}"`);
      let children = "";
      if (link) {
        attrs.push(` UpdateMode="${link.updateMode === "onCall" ? "OnCall" : "Always"}"`);
        const innerEls: string[] = [];
        if (link.lockedField !== undefined) {
          innerEls.push(`<o:LockedField>${link.lockedField ? "t" : "f"}</o:LockedField>`);
        }
        if (link.fieldCodes)
          innerEls.push(`<o:FieldCodes>${escapeXml(link.fieldCodes)}</o:FieldCodes>`);
        children = innerEls.join("");
      }
      const open = `<o:OLEObject${attrs.join("")}`;
      inner.push(children ? `${open}>${children}</o:OLEObject>` : `${open}/>`);
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

  parse(el, ctx) {
    const result: Partial<ObjectElementOptions> = {};

    const dxaOrig = attrNum(el, "w:dxaOrig");
    if (dxaOrig !== undefined) result.dxaOrig = dxaOrig;
    const dyaOrig = attrNum(el, "w:dyaOrig");
    if (dyaOrig !== undefined) result.dyaOrig = dyaOrig;

    // v:shapetype preamble (Word's _x0000_t75 OLE shapetype)
    const shapetypeEl = findChild(el, "v:shapetype");
    if (shapetypeEl) result.shapetype = parseVmlShapetype(shapetypeEl);

    // VML shape — structural capture; binary media is fetched through the
    // part's rels (r:id → media path → raw bytes) so round-trips keep the data.
    const shape = findChild(el, "v:shape");
    if (shape) {
      const id = attr(shape, "id");
      if (id) result.shapeId = id;
      const style = attr(shape, "style");
      if (style) {
        const parsed = parseVmlStyle(style);
        if (parsed["width"]) result.width = parsed["width"] as UniversalMeasure;
        if (parsed["height"]) result.height = parsed["height"] as UniversalMeasure;
      }
      const imagedataEl = findChild(shape, "v:imagedata");
      if (imagedataEl) {
        const imagedata = parseVmlImageData(imagedataEl);
        const media = resolveBinary(imagedata.relationshipId, ctx);
        result.iconImage = {
          data: media?.bytes ?? new Uint8Array(),
          type: media ? extensionOf(media.path) : "",
          ...(imagedata.officeTitle !== undefined ? { title: imagedata.officeTitle } : {}),
        };
      }
    }

    // Choice elements — o:OLEObject is what Word writes; the w:objectEmbed/
    // w:objectLink spellings are accepted for XSD-conforming third-party files.
    const oleEl = findChild(el, "o:OLEObject");
    if (oleEl) {
      const common = parseOleObject(oleEl);
      const payload = resolveBinary(attr(oleEl, "r:id"), ctx);
      if (payload) common.data = payload.bytes;
      if (attr(oleEl, "Type") === "Link") {
        const updateMode = attr(oleEl, "UpdateMode");
        const lockedFieldEl = findChild(oleEl, "o:LockedField");
        result.link = {
          ...common,
          updateMode: updateMode === "OnCall" ? "onCall" : "always",
          ...(lockedFieldEl ? { lockedField: parseOnOff(textOf(lockedFieldEl)) ?? false } : {}),
        };
      } else {
        result.embed = common;
      }
    }

    const embedEl = findChild(el, "w:objectEmbed");
    if (embedEl) {
      const embed = parseEmbed(embedEl);
      const payload = resolveBinary(attr(embedEl, "r:id"), ctx);
      if (payload) embed.data = payload.bytes;
      result.embed = embed;
    }

    const linkEl = findChild(el, "w:objectLink");
    if (linkEl) {
      const base = parseEmbed(linkEl);
      const payload = resolveBinary(attr(linkEl, "r:id"), ctx);
      if (payload) base.data = payload.bytes;
      const updateMode = attr(linkEl, "w:updateMode");
      const lockedField = attr(linkEl, "w:lockedField");
      result.link = {
        ...base,
        ...(updateMode ? { updateMode: updateMode as "always" | "onCall" } : {}),
        ...(lockedField !== undefined ? { lockedField: parseOnOff(lockedField) ?? false } : {}),
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

/** Extract the file extension from an image reference ("" when absent). */
function extensionOf(ref: string): string {
  const dot = ref.lastIndexOf(".");
  if (dot === -1) return "";
  return ref.slice(dot + 1);
}

/** Resolve a relationship id to its binary part bytes (media, OLE embedding). */
function resolveBinary(
  rId: string | undefined,
  ctx: ReadContext,
): { path: string; bytes: Uint8Array } | undefined {
  if (!rId) return undefined;
  const path = ctx.resolveRelationship(rId);
  if (!path) return undefined;
  const bytes = ctx.getRaw(path);
  if (!bytes) return undefined;
  return { path, bytes };
}

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

/** Parse common o:OLEObject attributes (excludes r:id — external on parse). */
function parseOleObject(el: Element): ObjectEmbedOptions {
  const opts: Partial<ObjectEmbedOptions> = {};
  const drawAspect = attr(el, "DrawAspect");
  if (drawAspect === "Icon") opts.drawAspect = "icon";
  else if (drawAspect === "Content") opts.drawAspect = "content";
  const progId = attr(el, "ProgID");
  if (progId) opts.progId = progId;
  const shapeId = attr(el, "ShapeID");
  if (shapeId) opts.shapeId = shapeId;
  const objectId = attr(el, "ObjectID");
  if (objectId) opts.objectId = objectId;
  const fieldCodesEl = findChild(el, "o:FieldCodes");
  if (fieldCodesEl) opts.fieldCodes = textOf(fieldCodesEl);
  // Placeholder until the caller resolves r:id against the part's rels.
  opts.data = new Uint8Array();
  return opts as ObjectEmbedOptions;
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
  // Placeholder until the caller resolves r:id against the part's rels.
  opts.data = new Uint8Array();
  return opts as ObjectEmbedOptions;
}
