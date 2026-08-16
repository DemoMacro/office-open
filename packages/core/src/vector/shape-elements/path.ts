/**
 * v:path element — CT_Path (path geometry, textbox rectangle, render toggles).
 *
 * Reference: ISO/IEC 29500-4, vml-main.xsd, CT_Path.
 *
 * @module
 */
import type { Element as XmlElement } from "@office-open/xml";

import {
  stringifyVmlAttributes,
  parseVmlAttributes,
  type VmlTrueFalse,
  type VmlAttrSpec,
} from "../attributes";
import type { VmlConnectType } from "./office-elements";

/** v:path options (CT_Path). */
export interface VmlPathOptions {
  id?: string;
  /** Path command string ("m0,0l100,0,100,100,0,100xe"). */
  v?: string;
  limo?: string;
  textboxrect?: string;
  fillok?: VmlTrueFalse;
  strokeok?: VmlTrueFalse;
  shadowok?: VmlTrueFalse;
  arrowok?: VmlTrueFalse;
  gradientshapeok?: VmlTrueFalse;
  textpathok?: VmlTrueFalse;
  insetpenok?: VmlTrueFalse;
  // o: extension members (AG_PathAttributes o: refs)
  /** o:connecttype — how shapes connect at this path. */
  connecttype?: VmlConnectType;
  /** o:connectlocs — connection-site coordinates. */
  connectlocs?: string;
  /** o:connectangles — connection-site angles. */
  connectangles?: string;
  /** o:extrusionok — 3D extrusion allowed along this path. */
  extrusionok?: VmlTrueFalse;
}

const PATH_ATTRS: readonly VmlAttrSpec[] = [
  { field: "id", attr: "id", kind: "string" },
  { field: "v", attr: "v", kind: "string" },
  { field: "limo", attr: "limo", kind: "string" },
  { field: "textboxrect", attr: "textboxrect", kind: "string" },
  { field: "fillok", attr: "fillok", kind: "trueFalse" },
  { field: "strokeok", attr: "strokeok", kind: "trueFalse" },
  { field: "shadowok", attr: "shadowok", kind: "trueFalse" },
  { field: "arrowok", attr: "arrowok", kind: "trueFalse" },
  { field: "gradientshapeok", attr: "gradientshapeok", kind: "trueFalse" },
  { field: "textpathok", attr: "textpathok", kind: "trueFalse" },
  { field: "insetpenok", attr: "insetpenok", kind: "trueFalse" },
  { field: "connecttype", attr: "o:connecttype", kind: "string" },
  { field: "connectlocs", attr: "o:connectlocs", kind: "string" },
  { field: "connectangles", attr: "o:connectangles", kind: "string" },
  { field: "extrusionok", attr: "o:extrusionok", kind: "trueFalse" },
];

/** Serialize v:path. */
export function stringifyVmlPath(opts: VmlPathOptions): string {
  return `<v:path${stringifyVmlAttributes(opts as Record<string, unknown>, PATH_ATTRS)}/>`;
}

/** Parse a v:path element. */
export function parseVmlPath(el: XmlElement): VmlPathOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, PATH_ATTRS, out);
  return out as VmlPathOptions;
}
