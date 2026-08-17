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
import type { ReadContext } from "@office-open/core/descriptor";
import type { Element } from "@office-open/xml";
import type { MediaData } from "@shared/media";

import type { BodyContext } from "../../context";

// ── Options ──

/** Binary backing one v:imagedata r:id reference in the shape tree. */
export interface PictMediaOptions {
  /** Placeholder key matching the `{fileName}` token in the imagedata r:id. */
  fileName: string;
  /** Image bytes captured from the source part. */
  data: Uint8Array;
  /** Image type / extension (e.g. "png", "wmf"). */
  type: string;
}

export interface PictOptions {
  /** Ordered shape elements — shapetype preambles, shapes, groups, … */
  children?: VmlShapeChild[];
  /** Binaries referenced by v:imagedata r:id (round-trip; authoring supplies
   *  media entries and matching `{fileName}` relationshipId placeholders). */
  media?: PictMediaOptions[];
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
  const children = opts.children ?? [];
  if (renames.size > 0) remapPlaceholders(children, renames);
  const inner = children.map(stringifyVmlShapeChild).join("");
  return inner !== "" ? `<w:pict>${inner}</w:pict>` : "<w:pict/>";
}

/** Replace `{old}` imagedata placeholders with their registered file names. */
function remapPlaceholders(children: VmlShapeChild[], renames: Map<string, string>): void {
  for (const child of children) {
    const fields = shapeFieldsOf(child);
    const rid = fields.imagedata?.relationshipId;
    if (rid !== undefined) {
      const mapped = renames.get(rid.slice(1, -1));
      if (mapped !== undefined) fields.imagedata!.relationshipId = `{${mapped}}`;
    }
    if ("group" in child) remapPlaceholders(child.group.children ?? [], renames);
  }
}

// ── Parse ──

/** Parse w:pict into ordered shape children, bridging imagedata r:id media. */
export function parsePict(el: Element, ctx: ReadContext): PictOptions {
  const children: VmlShapeChild[] = [];
  for (const child of el.elements ?? []) {
    if (child.type !== "element") continue;
    const shapeChild = parseVmlShapeChild(child);
    if (shapeChild !== undefined) children.push(shapeChild);
  }
  const media: PictMediaOptions[] = [];
  bridgeImagedata(children, ctx, media);
  return {
    ...(children.length > 0 ? { children } : {}),
    ...(media.length > 0 ? { media } : {}),
  };
}

/** Capture the binary behind each imagedata r:id, leaving a `{fileName}`
 *  placeholder for the stringify-side re-registration. */
function bridgeImagedata(
  children: VmlShapeChild[],
  ctx: ReadContext,
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
        const fileName = path.slice(path.lastIndexOf("/") + 1);
        media.push({ fileName, data: bytes, type: extensionOf(path) });
        img!.relationshipId = `{${fileName}}`;
      }
    }
    if ("group" in child) bridgeImagedata(child.group.children ?? [], ctx, media);
  }
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
