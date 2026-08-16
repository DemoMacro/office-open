/**
 * v:imagedata element — CT_ImageData.
 *
 * AG_ImageAttributes (crop/gain/gamma …) plus chromakey, the emboss/recolor
 * colors, and the three relationship references. Relationship ids stay plain
 * strings — callers that bridge media hand in the `{fileName}` placeholder
 * form.
 *
 * Reference: ISO/IEC 29500-4, vml-main.xsd, CT_ImageData / AG_ImageAttributes.
 *
 * @module
 */
import type { Element as XmlElement } from "@office-open/xml";

import {
  stringifyVmlAttributes,
  parseVmlAttributes,
  type VmlColor,
  type VmlTrueFalse,
  type VmlAttrSpec,
} from "../attributes";

/** v:imagedata options (CT_ImageData). */
export interface VmlImageDataOptions {
  id?: string;
  src?: string;
  cropleft?: string;
  croptop?: string;
  cropright?: string;
  cropbottom?: string;
  gain?: string;
  blacklevel?: string;
  gamma?: string;
  grayscale?: VmlTrueFalse;
  bilevel?: VmlTrueFalse;
  chromakey?: VmlColor;
  embosscolor?: VmlColor;
  recolortarget?: VmlColor;
  /** r:id — picture relationship; callers bridge media via the placeholder form. */
  relationshipId?: string;
  /** r:pict — print-output relationship. */
  pictRelationshipId?: string;
  /** r:href — hyperlink relationship. */
  hrefRelationshipId?: string;
  // o: extension members
  /** o:href — hyperlink target. */
  officeHref?: string;
  /** o:althref — alternate (high-contrast) image reference. */
  officeAltHref?: string;
  /** o:title — image title (the tooltip Office shows). */
  officeTitle?: string;
  /** o:oleid — OLE object number. */
  oleid?: number;
  /** o:detectmouseclick — click detection for OLE activation. */
  detectmouseclick?: VmlTrueFalse;
  /** o:movie — movie reference number. */
  movie?: number;
  /** o:relid — relationship id, in the o: extension slot. */
  officeRelationshipId?: string;
}

const IMAGEDATA_ATTRS: readonly VmlAttrSpec[] = [
  { field: "id", attr: "id", kind: "string" },
  { field: "src", attr: "src", kind: "string" },
  { field: "cropleft", attr: "cropleft", kind: "string" },
  { field: "croptop", attr: "croptop", kind: "string" },
  { field: "cropright", attr: "cropright", kind: "string" },
  { field: "cropbottom", attr: "cropbottom", kind: "string" },
  { field: "gain", attr: "gain", kind: "string" },
  { field: "blacklevel", attr: "blacklevel", kind: "string" },
  { field: "gamma", attr: "gamma", kind: "string" },
  { field: "grayscale", attr: "grayscale", kind: "trueFalse" },
  { field: "bilevel", attr: "bilevel", kind: "trueFalse" },
  { field: "chromakey", attr: "chromakey", kind: "string" },
  { field: "embosscolor", attr: "embosscolor", kind: "string" },
  { field: "recolortarget", attr: "recolortarget", kind: "string" },
  { field: "relationshipId", attr: "r:id", kind: "string" },
  { field: "pictRelationshipId", attr: "r:pict", kind: "string" },
  { field: "hrefRelationshipId", attr: "r:href", kind: "string" },
  { field: "officeHref", attr: "o:href", kind: "string" },
  { field: "officeAltHref", attr: "o:althref", kind: "string" },
  { field: "officeTitle", attr: "o:title", kind: "string" },
  { field: "oleid", attr: "o:oleid", kind: "number" },
  { field: "detectmouseclick", attr: "o:detectmouseclick", kind: "trueFalse" },
  { field: "movie", attr: "o:movie", kind: "number" },
  { field: "officeRelationshipId", attr: "o:relid", kind: "string" },
];

/** Serialize v:imagedata. */
export function stringifyVmlImageData(opts: VmlImageDataOptions): string {
  return `<v:imagedata${stringifyVmlAttributes(opts as Record<string, unknown>, IMAGEDATA_ATTRS)}/>`;
}

/** Parse a v:imagedata element. */
export function parseVmlImageData(el: XmlElement): VmlImageDataOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, IMAGEDATA_ATTRS, out);
  return out as VmlImageDataOptions;
}
