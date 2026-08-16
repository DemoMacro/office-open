/**
 * v:fill element — CT_Fill.
 *
 * Solid / gradient / tile / pattern / frame fill for VML shapes; also the
 * child of v:background. Relationship references (r:id) stay plain strings —
 * callers that bridge media hand in the `{fileName}` placeholder form, exactly
 * like the DrawingML blip path.
 *
 * Reference: ISO/IEC 29500-4, vml-main.xsd, CT_Fill / ST_FillType /
 * ST_FillMethod.
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
import {
  stringifyVmlOfficeFill,
  parseVmlOfficeFill,
  type VmlOfficeFillOptions,
} from "./office-elements";

/** ST_FillType. */
export type VmlFillType = "solid" | "gradient" | "gradientRadial" | "tile" | "pattern" | "frame";

/** ST_FillMethod. */
export type VmlFillMethod = "none" | "linear" | "sigma" | "any" | "linear sigma";

/** ST_ImageAspect. */
export type VmlImageAspect = "ignore" | "atMost" | "atLeast";

/** v:fill options (CT_Fill). */
export interface VmlFillOptions {
  id?: string;
  type?: VmlFillType;
  on?: VmlTrueFalse;
  color?: VmlColor;
  opacity?: string;
  color2?: VmlColor;
  src?: string;
  size?: string;
  origin?: string;
  position?: string;
  aspect?: VmlImageAspect;
  colors?: string;
  /** Gradient angle in degrees. */
  angle?: number;
  alignshape?: VmlTrueFalse;
  focus?: string;
  focussize?: string;
  focusposition?: string;
  method?: VmlFillMethod;
  recolor?: VmlTrueFalse;
  rotate?: VmlTrueFalse;
  /** Relationship reference; callers bridge media via the placeholder form. */
  relationshipId?: string;
  // o: extension members (AG_FillAttributes o: refs)
  /** o:href — hyperlink target for image fills. */
  officeHref?: string;
  /** o:althref — alternate (high-contrast) image reference. */
  officeAltHref?: string;
  /** o:detectmouseclick — the fill toggles on mouse click. */
  detectmouseclick?: VmlTrueFalse;
  /** o:title — image title. */
  officeTitle?: string;
  /** o:opacity2 — the second-stop opacity. */
  opacity2?: string;
  /** o:relid — relationship id, in the o: extension slot. */
  officeRelationshipId?: string;
  /** o:fill child — the office fill-type extension. */
  officeFill?: VmlOfficeFillOptions;
}

const FILL_ATTRS: readonly VmlAttrSpec[] = [
  { field: "id", attr: "id", kind: "string" },
  { field: "type", attr: "type", kind: "string" },
  { field: "on", attr: "on", kind: "trueFalse" },
  { field: "color", attr: "color", kind: "string" },
  { field: "opacity", attr: "opacity", kind: "string" },
  { field: "color2", attr: "color2", kind: "string" },
  { field: "src", attr: "src", kind: "string" },
  { field: "size", attr: "size", kind: "string" },
  { field: "origin", attr: "origin", kind: "string" },
  { field: "position", attr: "position", kind: "string" },
  { field: "aspect", attr: "aspect", kind: "string" },
  { field: "colors", attr: "colors", kind: "string" },
  { field: "angle", attr: "angle", kind: "number" },
  { field: "alignshape", attr: "alignshape", kind: "trueFalse" },
  { field: "focus", attr: "focus", kind: "string" },
  { field: "focussize", attr: "focussize", kind: "string" },
  { field: "focusposition", attr: "focusposition", kind: "string" },
  { field: "method", attr: "method", kind: "string" },
  { field: "recolor", attr: "recolor", kind: "trueFalse" },
  { field: "rotate", attr: "rotate", kind: "trueFalse" },
  { field: "relationshipId", attr: "r:id", kind: "string" },
  { field: "officeHref", attr: "o:href", kind: "string" },
  { field: "officeAltHref", attr: "o:althref", kind: "string" },
  { field: "detectmouseclick", attr: "o:detectmouseclick", kind: "trueFalse" },
  { field: "officeTitle", attr: "o:title", kind: "string" },
  { field: "opacity2", attr: "o:opacity2", kind: "string" },
  { field: "officeRelationshipId", attr: "o:relid", kind: "string" },
];

/** Serialize v:fill. */
export function stringifyVmlFill(opts: VmlFillOptions): string {
  const child = opts.officeFill !== undefined ? stringifyVmlOfficeFill(opts.officeFill) : "";
  const attrStr = stringifyVmlAttributes(opts as Record<string, unknown>, FILL_ATTRS);
  return child !== "" ? `<v:fill${attrStr}>${child}</v:fill>` : `<v:fill${attrStr}/>`;
}

/** Parse a v:fill element. */
export function parseVmlFill(el: XmlElement): VmlFillOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, FILL_ATTRS, out);
  for (const child of el.elements ?? []) {
    if (child.type === "element" && child.name === "o:fill") {
      out.officeFill = parseVmlOfficeFill(child);
      break;
    }
  }
  return out as VmlFillOptions;
}
