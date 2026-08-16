/**
 * v:shadow element — CT_Shadow (single / double / emboss / perspective).
 *
 * Reference: ISO/IEC 29500-4, vml-main.xsd, CT_Shadow / ST_ShadowType.
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

/** ST_ShadowType. */
export type VmlShadowType = "single" | "double" | "emboss" | "perspective";

/** v:shadow options (CT_Shadow). */
export interface VmlShadowOptions {
  id?: string;
  on?: VmlTrueFalse;
  type?: VmlShadowType;
  obscured?: VmlTrueFalse;
  color?: VmlColor;
  opacity?: string;
  offset?: string;
  color2?: VmlColor;
  offset2?: string;
  origin?: string;
  matrix?: string;
}

const SHADOW_ATTRS: readonly VmlAttrSpec[] = [
  { field: "id", attr: "id", kind: "string" },
  { field: "on", attr: "on", kind: "trueFalse" },
  { field: "type", attr: "type", kind: "string" },
  { field: "obscured", attr: "obscured", kind: "trueFalse" },
  { field: "color", attr: "color", kind: "string" },
  { field: "opacity", attr: "opacity", kind: "string" },
  { field: "offset", attr: "offset", kind: "string" },
  { field: "color2", attr: "color2", kind: "string" },
  { field: "offset2", attr: "offset2", kind: "string" },
  { field: "origin", attr: "origin", kind: "string" },
  { field: "matrix", attr: "matrix", kind: "string" },
];

/** Serialize v:shadow. */
export function stringifyVmlShadow(opts: VmlShadowOptions): string {
  return `<v:shadow${stringifyVmlAttributes(opts as Record<string, unknown>, SHADOW_ATTRS)}/>`;
}

/** Parse a v:shadow element. */
export function parseVmlShadow(el: XmlElement): VmlShadowOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, SHADOW_ATTRS, out);
  return out as VmlShadowOptions;
}
