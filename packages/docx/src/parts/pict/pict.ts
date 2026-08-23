/**
 * VML picture element — w:pict.
 *
 * A run-level pre-DrawingML drawing object (CT_Picture is xsd:any over the VML
 * shape vocabulary). Children round-trip as ordered shape wrappers; v:imagedata
 * binaries are captured through the part's rels on parse and re-registered on
 * stringify, with r:id carried as `{fileName}` placeholders the compiler's
 * media bridge rewrites.
 *
 * Reference: OOXML transitional, wml.xsd CT_Picture; vml-main.xsd shapes.
 *
 * @module
 */
import { stringifyVmlShapeChild, parseVmlShapeChild } from "@office-open/core";
import type { VmlBaseShapeFields, VmlShapeChild } from "@office-open/core";
import type { Element } from "@office-open/xml";
import type { MediaData } from "@shared/media";

import type { BodyContext, DocxReadContext } from "../../context";
import { replaceRelsWithPlaceholders } from "../../util/replace-media-placeholders";

// ── Options ──

/** Binary backing one `{fileName}` placeholder in the pict (imagedata or raw
 *  passthrough XML — fallback copies, textbox content). */
export interface PictMediaOptions {
  /** Placeholder key matching the `{fileName}` token in the pict XML. */
  fileName: string;
  /** Image bytes captured from the source part. */
  data: Uint8Array;
  /** Image type / extension (e.g. "png", "wmf"). */
  type: string;
}

export interface PictOptions {
  /** Ordered shape elements — shapetype preambles, shapes, groups, … */
  children?: VmlShapeChild[];
  /** Binaries referenced by `{fileName}` placeholders — v:imagedata r:id in
   *  the shape tree plus any r:id/r:embed/r:link inside the VML fallback or
   *  textbox content (round-trip; authoring supplies media entries and
   *  matching placeholders). */
  media?: PictMediaOptions[];
  /** Serialized mc:Fallback element carried when the source wrapped this pict
   *  in mc:AlternateContent. Round-trips verbatim (wrapper rebuilt on
   *  stringify); round-trip only — do not hand-author. */
  vmlFallback?: string;
  /** mc:Choice @Requires namespace prefix; defaults to "w14" on stringify. */
  mcChoiceRequires?: string;
  /** w14:anchorId extension attribute on w:pict. */
  w14AnchorId?: string;
}

// ── Stringify ──

/**
 * Serialize w:pict: register imagedata media (dedup may rename — the shape
 * tree's `{fileName}` placeholders are remapped to the registered names), then
 * emit the ordered children.
 *
 * The context is narrowed to the media registry — callers outside the body
 * descriptor chain (e.g. numbering picture bullets) pass `{ file: writeCtx }`.
 */
export function stringifyPict(opts: PictOptions, ctx: Pick<BodyContext, "file">): string {
  const renames = new Map<string, string>();
  for (const m of opts.media ?? []) {
    // Skip empty/extensionless entries — they would register a 0-byte part
    // with no covering [Content_Types] entry (an OPC violation).
    if (!(m.data.length > 0 && m.type)) continue;
    const data = m.data;
    const type = m.type;
    const { fileName } = ctx.file.media.addMedia(
      data,
      type,
      // Binary-only registration — the pict renders via VML, so the drawing
      // transformation is unused and left zeroed (vmlFallback media pattern).
      (fileName) =>
        ({
          type,
          data,
          fileName,
          transformation: { emus: { x: 0, y: 0 }, pixels: { x: 0, y: 0 } },
        }) as MediaData,
      m.fileName,
    );
    if (fileName !== m.fileName) renames.set(m.fileName, fileName);
  }
  let children = opts.children ?? [];
  let vmlFallback = opts.vmlFallback;
  if (renames.size > 0) {
    // Remap a copy — the caller's PictOptions must stay untouched (options
    // objects are shared, serializable data; mutating them would leak renamed
    // placeholders into the next document built from the same object).
    children = structuredClone(children);
    remapPlaceholders(children, renames);
    if (vmlFallback !== undefined) vmlFallback = remapRawPlaceholders(vmlFallback, renames);
  }
  const anchor = opts.w14AnchorId !== undefined ? ` w14:anchorId="${opts.w14AnchorId}"` : "";
  const inner = children.map(stringifyVmlShapeChild).join("");
  const pictXml = inner !== "" ? `<w:pict${anchor}>${inner}</w:pict>` : `<w:pict${anchor}/>`;
  if (vmlFallback !== undefined) {
    return `<mc:AlternateContent><mc:Choice Requires="${opts.mcChoiceRequires ?? "w14"}">${pictXml}</mc:Choice>${vmlFallback}</mc:AlternateContent>`;
  }
  return pictXml;
}

/** Replace `{old}` imagedata/textbox placeholders with their registered names. */
function remapPlaceholders(children: VmlShapeChild[], renames: Map<string, string>): void {
  for (const child of children) {
    const fields = shapeFieldsOf(child);
    const rid = fields.imagedata?.relationshipId;
    if (rid !== undefined) {
      const mapped = renames.get(rid.slice(1, -1));
      if (mapped !== undefined) fields.imagedata!.relationshipId = `{${mapped}}`;
    }
    const textbox = "shape" in child ? child.shape.textbox : undefined;
    if (textbox?.txbxContent !== undefined) {
      textbox.txbxContent = remapRawPlaceholders(textbox.txbxContent, renames);
    }
    if ("group" in child) remapPlaceholders(child.group.children ?? [], renames);
  }
}

/** Replace `{old}` placeholders inside a raw XML string. */
function remapRawPlaceholders(xml: string, renames: Map<string, string>): string {
  for (const [oldName, newName] of renames) {
    xml = xml.split(`{${oldName}}`).join(`{${newName}}`);
  }
  return xml;
}

// ── Parse ──

/** Parse w:pict into ordered shape children, bridging imagedata r:id media. */
export function parsePict(el: Element, ctx: DocxReadContext): PictOptions {
  const children: VmlShapeChild[] = [];
  for (const child of el.elements ?? []) {
    if (child.type !== "element") continue;
    const shapeChild = parseVmlShapeChild(child);
    if (shapeChild !== undefined) children.push(shapeChild);
  }
  const media: PictMediaOptions[] = [];
  bridgeImagedata(children, ctx, media);
  bridgeTxbxContent(children, ctx, media);
  return {
    ...(children.length > 0 ? { children } : {}),
    ...(media.length > 0 ? { media } : {}),
  };
}

/**
 * Replace relationship references inside a raw XML string (textbox content,
 * mc:Fallback copies) with `{fileName}` placeholders and collect the media.
 *
 * Raw passthrough keeps the source rIds verbatim, but the generated rels
 * renumber — the placeholders let the compiler's media bridge re-register the
 * binaries and resolve fresh ids (same pattern as v:imagedata).
 */
function bridgeRawRels(xml: string, ctx: DocxReadContext, media: PictMediaOptions[]): string {
  const { rawXml, rawMedia } = replaceRelsWithPlaceholders(xml, ctx);
  for (const m of rawMedia) {
    if (!media.some((e) => e.fileName === m.fileName)) {
      media.push({ fileName: m.fileName, data: m.data as Uint8Array, type: m.type });
    }
  }
  return rawXml;
}

/** Bridge r:id/r:embed/r:link references inside each shape's textbox content. */
function bridgeTxbxContent(
  children: VmlShapeChild[],
  ctx: DocxReadContext,
  media: PictMediaOptions[],
): void {
  for (const child of children) {
    const textbox = "shape" in child ? child.shape.textbox : undefined;
    if (textbox?.txbxContent !== undefined) {
      textbox.txbxContent = bridgeRawRels(textbox.txbxContent, ctx, media);
    }
    if ("group" in child) bridgeTxbxContent(child.group.children ?? [], ctx, media);
  }
}

/** Capture the binary behind each imagedata r:id, leaving a `{fileName}`
 *  placeholder for the stringify-side re-registration. */
function bridgeImagedata(
  children: VmlShapeChild[],
  ctx: DocxReadContext,
  media: PictMediaOptions[],
): void {
  for (const child of children) {
    const fields = shapeFieldsOf(child);
    const img = fields.imagedata;
    const rid = img?.relationshipId;
    if (rid !== undefined) {
      const path = ctx.resolveRelationship(rid);
      const bytes = path ? ctx.getRaw(path) : undefined;
      // Unresolvable references (dangling rel, empty part) stay verbatim —
      // registering a nameless 0-byte media part would violate OPC.
      if (path && bytes && bytes.length > 0) {
        let fileName = path.slice(path.lastIndexOf("/") + 1);
        // Placeholders key by file name — two rels pointing at different bytes
        // under the same basename must not collapse onto one placeholder.
        fileName = uniqueFileName(media, fileName);
        media.push({ fileName, data: bytes, type: extensionOf(path) });
        img!.relationshipId = `{${fileName}}`;
      }
    }
    if ("group" in child) bridgeImagedata(child.group.children ?? [], ctx, media);
  }
}

/** Disambiguate a media file name already taken within one pict. */
function uniqueFileName(media: PictMediaOptions[], fileName: string): string {
  const taken = new Set(media.map((m) => m.fileName));
  if (!taken.has(fileName)) return fileName;
  const dot = fileName.lastIndexOf(".");
  const stem = dot === -1 ? fileName : fileName.slice(0, dot);
  const ext = dot === -1 ? "" : fileName.slice(dot);
  let i = 2;
  while (taken.has(`${stem}-${i}${ext}`)) i++;
  return `${stem}-${i}${ext}`;
}

/** The shape payload of a wrapper — every variant extends VmlBaseShapeFields. */
function shapeFieldsOf(child: VmlShapeChild): VmlBaseShapeFields {
  if ("shape" in child) return child.shape;
  if ("shapetype" in child) return child.shapetype;
  if ("group" in child) return child.group;
  if ("arc" in child) return child.arc;
  if ("curve" in child) return child.curve;
  if ("image" in child) return child.image;
  if ("line" in child) return child.line;
  if ("oval" in child) return child.oval;
  if ("polyline" in child) return child.polyline;
  if ("rect" in child) return child.rect;
  return child.roundrect;
}

/** Extract the file extension from a part path ("" when absent). */
function extensionOf(ref: string): string {
  const dot = ref.lastIndexOf(".");
  if (dot === -1) return "";
  return ref.slice(dot + 1);
}
