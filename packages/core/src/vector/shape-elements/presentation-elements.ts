/**
 * Presentation pvml: elements — the two EG_ShapeElements members from
 * vml-presentationDrawing.xsd: iscomment (empty marker) and textdata.
 *
 * Reference: ISO/IEC 29500-4, vml-presentationDrawing.xsd.
 *
 * @module
 */
import type { Element as XmlElement } from "@office-open/xml";

/** pvml:iscomment options (CT_Empty) — empty marker element. */
export interface VmlIsCommentOptions {}

/** Serialize pvml:iscomment. */
export function stringifyVmlIsComment(_opts: VmlIsCommentOptions): string {
  return "<pvml:iscomment/>";
}

/** Parse a pvml:iscomment element. */
export function parseVmlIsComment(_el: XmlElement): VmlIsCommentOptions {
  return {};
}

/** pvml:textdata options (CT_Rel). */
export interface VmlTextDataOptions {
  /** Relationship id — bridged by the caller. */
  id?: string;
}

/** Serialize pvml:textdata. */
export function stringifyVmlTextData(opts: VmlTextDataOptions): string {
  return opts.id !== undefined ? `<pvml:textdata id="${opts.id}"/>` : "<pvml:textdata/>";
}

/** Parse a pvml:textdata element. */
export function parseVmlTextData(el: XmlElement): VmlTextDataOptions {
  const id = el.attributes?.id;
  return id !== undefined ? { id: String(id) } : {};
}
