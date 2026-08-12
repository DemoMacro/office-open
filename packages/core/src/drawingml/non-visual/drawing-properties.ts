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

/** Mirrors a:CT_NonVisualDrawingProps — name/descr/title/hidden. id is runtime; hyperlinks stay package-side. */
export interface NonVisualDrawingPropertiesOptions {
  /** XSD @name (required in XML); API-optional, the descriptor synthesizes a fallback. */
  name?: string;
  /** XSD @descr — alt text for accessibility; emitted only when present. */
  description?: string;
  /** XSD @title. */
  title?: string;
  /** XSD @hidden (default false); emitted as "1" only when true. */
  hidden?: boolean;
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
  return innerXml ? `<${tag} ${attrs}>${innerXml}</${tag}>` : `<${tag} ${attrs}/>`;
}

/**
 * Parse cNvPr/docPr attributes into NonVisualDrawingPropertiesOptions.
 * Empty descr/title are dropped (Word never round-trips them empty).
 */
export function parseNonVisualDrawingProperties(
  el: XmlElement | undefined,
): Partial<NonVisualDrawingPropertiesOptions> {
  if (!el?.attributes) return {};
  const a = el.attributes;
  const result: Partial<NonVisualDrawingPropertiesOptions> = {};
  if (a["name"] !== undefined) result.name = String(a["name"]);
  const descr = a["descr"];
  if (descr !== undefined && descr !== "") result.description = String(descr);
  const title = a["title"];
  if (title !== undefined && title !== "") result.title = String(title);
  const hidden = a["hidden"];
  if (hidden !== undefined) result.hidden = hidden === "1" || hidden === "true";
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
  return picked;
}
