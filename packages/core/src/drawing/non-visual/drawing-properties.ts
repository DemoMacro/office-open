/**
 * Non-visual drawing properties (a:CT_NonVisualDrawingProps).
 *
 * The shared XSD type behind every drawing object's cNvPr (pic:/p:/xdr:/a:) and
 * docx's wp:docPr — picture, shape, connector, group, graphicFrame, chart, ole,
 * smartart each carry one. This module is the single source of truth for its
 * four user-facing attributes (name/descr/title/hidden). The runtime id is
 * allocated by each descriptor's nextId counter, and hlinkClick/hlinkHover stay
 * package-side (relationship wiring differs per format).
 *
 * @module
 */

import { escapeXml } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { attr, findChild, stringifyElement } from "@office-open/xml";

import { extUriMatches } from "../../util/ext-uri";
import { parseOnOff } from "../../util/values";
import type { Guid } from "../../util/values";

/** The a16:creationId extension uri (CT_NonVisualDrawingProps extLst). */
const CREATION_ID_EXT_URI = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}";

/** Mirrors a:CT_NonVisualDrawingProps — name/descr/title/hidden. id is runtime; hyperlinks stay package-side. */
export interface NonVisualDrawingPropertiesOptions {
  /** XSD `@name` (required in XML); API-optional, the descriptor synthesizes a fallback. */
  name?: string;
  /** XSD `@descr` — alt text for accessibility; emitted only when present. */
  description?: string;
  /** XSD `@title`. */
  title?: string;
  /** XSD `@hidden` (default false); emitted as "1" only when true. */
  hidden?: boolean;
  /**
   * a16:creationId `@id` from the cNvPr/docPr extension list — Office's
   * per-object creation stamp, the sole known cNvPr ext content.
   */
  creationId?: Guid;
  /**
   * Verbatim `a:extLst` inner XML for extensions beyond creationId (a14
   * picture effects, useLocalDpi, …). Round-trip only: takes precedence over
   * `creationId`, which it subsumes when the source list carries both.
   */
  ext?: string;
}

/**
 * Stringify a cNvPr/docPr opening tag (a:CT_NonVisualDrawingProps attributes).
 *
 * `tag` is the fully-qualified element name ("p:cNvPr", "pic:cNvPr", "wp:docPr",
 * …) since the same XSD type is serialized under several element names across
 * the format packages. `innerXml` carries caller-built hlinkClick/hlinkHover.
 */
export function stringifyNonVisualDrawingProperties(
  tag: string,
  id: number | string,
  opts: NonVisualDrawingPropertiesOptions | undefined,
  fallbackName: string,
  innerXml?: string,
): string {
  const name = opts?.name ?? fallbackName;
  let attrs = `id="${id}" name="${escapeXml(name)}"`;
  if (opts?.description) attrs += ` descr="${escapeXml(opts.description)}"`;
  if (opts?.title) attrs += ` title="${escapeXml(opts.title)}"`;
  if (opts?.hidden) attrs += ` hidden="1"`;
  // CT_NonVisualDrawingProps tail: hlinkClick/hover (caller innerXml) → extLst.
  // xmlns:a is declared locally: parts that host docPr without a DrawingML
  // root (comments.xml, notes) never declare the a: prefix themselves.
  const A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
  const extLst = opts?.ext
    ? `<a:extLst ${A_NS}>${opts.ext}</a:extLst>`
    : opts?.creationId
      ? `<a:extLst ${A_NS}><a:ext uri="${CREATION_ID_EXT_URI}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="${escapeXml(opts.creationId)}"/></a:ext></a:extLst>`
      : "";
  const content = (innerXml ?? "") + extLst;
  return content ? `<${tag} ${attrs}>${content}</${tag}>` : `<${tag} ${attrs}/>`;
}

/**
 * Parse cNvPr/docPr attributes into NonVisualDrawingPropertiesOptions.
 * Empty descr/title are dropped (Word never round-trips them empty).
 */
export function parseNonVisualDrawingProperties(
  el: XmlElement | undefined,
): Partial<NonVisualDrawingPropertiesOptions> {
  if (!el) return {};
  const result: Partial<NonVisualDrawingPropertiesOptions> = {};
  if (el.attributes) {
    const a = el.attributes;
    if (a["name"] !== undefined) result.name = String(a["name"]);
    const descr = a["descr"];
    if (descr !== undefined && descr !== "") result.description = String(descr);
    const title = a["title"];
    if (title !== undefined && title !== "") result.title = String(title);
    const hidden = a["hidden"];
    if (hidden !== undefined) result.hidden = parseOnOff(hidden) ?? false;
  }
  const extLst = findChild(el, "a:extLst");
  if (extLst) {
    // Extensions beyond creationId (a14 picture effects, useLocalDpi, …) have
    // no structured model — keep the whole list verbatim and skip the
    // creationId extraction so the two channels never double-emit.
    const hasUnmodeled = (extLst.elements ?? []).some(
      (ext) => ext.name === "a:ext" && !extUriMatches(attr(ext, "uri"), CREATION_ID_EXT_URI),
    );
    if (hasUnmodeled) {
      const inner = (extLst.elements ?? []).map((e) => stringifyElement(e)).join("");
      if (inner) result.ext = inner;
    } else {
      for (const ext of extLst.elements ?? []) {
        if (ext.name !== "a:ext" || !extUriMatches(attr(ext, "uri"), CREATION_ID_EXT_URI)) continue;
        const creationId = findChild(ext, "a16:creationId");
        // a16:creationId keys its GUID on @id (p14:creationId uses @val).
        const id = creationId ? attr(creationId, "id") : undefined;
        if (id) result.creationId = id;
      }
    }
  }
  return result;
}

/**
 * Pick the cNvPr attributes (name/description/title/hidden) actually set on
 * `opts`, dropping undefined. Used when bridging a package's options onto a
 * descriptor that carries the same fields — spreads only what was authored so
 * the descriptor's fallbacks (e.g. synthesized name) still apply for the rest.
 */
export function pickNonVisualDrawingProperties(
  opts: NonVisualDrawingPropertiesOptions | undefined,
): Partial<NonVisualDrawingPropertiesOptions> {
  if (!opts) return {};
  const picked: Partial<NonVisualDrawingPropertiesOptions> = {};
  if (opts.name !== undefined) picked.name = opts.name;
  if (opts.description !== undefined) picked.description = opts.description;
  if (opts.title !== undefined) picked.title = opts.title;
  if (opts.hidden !== undefined) picked.hidden = opts.hidden;
  if (opts.creationId !== undefined) picked.creationId = opts.creationId;
  if (opts.ext !== undefined) picked.ext = opts.ext;
  return picked;
}
