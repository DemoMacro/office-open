/**
 * v:stroke element — CT_Stroke.
 *
 * AG_StrokeAttributes (~30 attributes, including the o: extension members)
 * plus the five o: border sub-strokes (o:left/top/right/bottom/column,
 * CT_StrokeChild).
 *
 * Reference: ISO/IEC 29500-4, vml-main.xsd, CT_Stroke / AG_StrokeAttributes.
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
import type { VmlFillType, VmlImageAspect } from "./fill";

/** ST_StrokeLineStyle. */
export type VmlStrokeLineStyle =
  | "single"
  | "thinThin"
  | "thinThick"
  | "thickThin"
  | "thickBetweenThin";

/** ST_StrokeJoinStyle. */
export type VmlStrokeJoinStyle = "round" | "bevel" | "miter";

/** ST_StrokeEndCap. */
export type VmlStrokeEndCap = "flat" | "square" | "round";

/** ST_StrokeArrowType. */
export type VmlStrokeArrowType = "none" | "block" | "classic" | "oval" | "diamond" | "open";

/** ST_StrokeArrowWidth. */
export type VmlStrokeArrowWidth = "narrow" | "medium" | "wide";

/** ST_StrokeArrowLength. */
export type VmlStrokeArrowLength = "short" | "medium" | "long";

/** v:stroke options (CT_Stroke / AG_StrokeAttributes). */
export interface VmlStrokeOptions {
  id?: string;
  on?: VmlTrueFalse;
  /** Line weight, e.g. "2pt". */
  weight?: string;
  color?: VmlColor;
  opacity?: string;
  linestyle?: VmlStrokeLineStyle;
  miterlimit?: number;
  joinstyle?: VmlStrokeJoinStyle;
  endcap?: VmlStrokeEndCap;
  /** Preset ("dash", "dotdotdash", …) or custom dash pattern ("4 2 4 2"). */
  dashstyle?: string;
  filltype?: VmlFillType;
  src?: string;
  imageaspect?: VmlImageAspect;
  imagesize?: string;
  imagealignshape?: VmlTrueFalse;
  color2?: VmlColor;
  startarrow?: VmlStrokeArrowType;
  startarrowwidth?: VmlStrokeArrowWidth;
  startarrowlength?: VmlStrokeArrowLength;
  endarrow?: VmlStrokeArrowType;
  endarrowwidth?: VmlStrokeArrowWidth;
  endarrowlength?: VmlStrokeArrowLength;
  /** Relationship reference; callers bridge media via the placeholder form. */
  relationshipId?: string;
  insetpen?: VmlTrueFalse;
  // AG_StrokeAttributes o: members
  /** o:href — hyperlink target for image fills. */
  officeHref?: string;
  /** o:althref — alternate (high-contrast) image reference. */
  officeAltHref?: string;
  /** o:title — image title. */
  officeTitle?: string;
  /** o:forcedash — always render dashed (accessibility). */
  forcedash?: VmlTrueFalse;
  /** o:relid — relationship id for the image, in the o: extension slot. */
  officeRelationshipId?: string;
  // o: border sub-strokes (CT_StrokeChild children)
  leftStroke?: VmlStrokeChildOptions;
  topStroke?: VmlStrokeChildOptions;
  rightStroke?: VmlStrokeChildOptions;
  bottomStroke?: VmlStrokeChildOptions;
  columnStroke?: VmlStrokeChildOptions;
}

/**
 * o:left/top/right/bottom/column options (CT_StrokeChild) — the per-edge
 * sub-strokes of v:stroke, sharing most of the stroke vocabulary.
 */
export interface VmlStrokeChildOptions {
  ext?: string;
  on?: VmlTrueFalse;
  weight?: string;
  color?: VmlColor;
  color2?: VmlColor;
  opacity?: string;
  linestyle?: VmlStrokeLineStyle;
  miterlimit?: number;
  joinstyle?: VmlStrokeJoinStyle;
  endcap?: VmlStrokeEndCap;
  dashstyle?: string;
  insetpen?: VmlTrueFalse;
  filltype?: VmlFillType;
  src?: string;
  imageaspect?: VmlImageAspect;
  imagesize?: string;
  imagealignshape?: VmlTrueFalse;
  startarrow?: VmlStrokeArrowType;
  startarrowwidth?: VmlStrokeArrowWidth;
  startarrowlength?: VmlStrokeArrowLength;
  endarrow?: VmlStrokeArrowType;
  endarrowwidth?: VmlStrokeArrowWidth;
  endarrowlength?: VmlStrokeArrowLength;
  /** o:href. */
  officeHref?: string;
  /** o:althref. */
  officeAltHref?: string;
  /** o:title. */
  officeTitle?: string;
  /** o:forcedash. */
  forcedash?: VmlTrueFalse;
}

const STROKE_ATTRS: readonly VmlAttrSpec[] = [
  { field: "id", attr: "id", kind: "string" },
  { field: "on", attr: "on", kind: "trueFalse" },
  { field: "weight", attr: "weight", kind: "string" },
  { field: "color", attr: "color", kind: "string" },
  { field: "opacity", attr: "opacity", kind: "string" },
  { field: "linestyle", attr: "linestyle", kind: "string" },
  { field: "miterlimit", attr: "miterlimit", kind: "number" },
  { field: "joinstyle", attr: "joinstyle", kind: "string" },
  { field: "endcap", attr: "endcap", kind: "string" },
  { field: "dashstyle", attr: "dashstyle", kind: "string" },
  { field: "filltype", attr: "filltype", kind: "string" },
  { field: "src", attr: "src", kind: "string" },
  { field: "imageaspect", attr: "imageaspect", kind: "string" },
  { field: "imagesize", attr: "imagesize", kind: "string" },
  { field: "imagealignshape", attr: "imagealignshape", kind: "trueFalse" },
  { field: "color2", attr: "color2", kind: "string" },
  { field: "startarrow", attr: "startarrow", kind: "string" },
  { field: "startarrowwidth", attr: "startarrowwidth", kind: "string" },
  { field: "startarrowlength", attr: "startarrowlength", kind: "string" },
  { field: "endarrow", attr: "endarrow", kind: "string" },
  { field: "endarrowwidth", attr: "endarrowwidth", kind: "string" },
  { field: "endarrowlength", attr: "endarrowlength", kind: "string" },
  { field: "relationshipId", attr: "r:id", kind: "string" },
  { field: "insetpen", attr: "insetpen", kind: "trueFalse" },
  { field: "officeHref", attr: "o:href", kind: "string" },
  { field: "officeAltHref", attr: "o:althref", kind: "string" },
  { field: "officeTitle", attr: "o:title", kind: "string" },
  { field: "forcedash", attr: "o:forcedash", kind: "trueFalse" },
  { field: "officeRelationshipId", attr: "o:relid", kind: "string" },
];

const STROKE_CHILD_ATTRS: readonly VmlAttrSpec[] = [
  { field: "ext", attr: "v:ext", kind: "string" },
  { field: "on", attr: "on", kind: "trueFalse" },
  { field: "weight", attr: "weight", kind: "string" },
  { field: "color", attr: "color", kind: "string" },
  { field: "color2", attr: "color2", kind: "string" },
  { field: "opacity", attr: "opacity", kind: "string" },
  { field: "linestyle", attr: "linestyle", kind: "string" },
  { field: "miterlimit", attr: "miterlimit", kind: "number" },
  { field: "joinstyle", attr: "joinstyle", kind: "string" },
  { field: "endcap", attr: "endcap", kind: "string" },
  { field: "dashstyle", attr: "dashstyle", kind: "string" },
  { field: "insetpen", attr: "insetpen", kind: "trueFalse" },
  { field: "filltype", attr: "filltype", kind: "string" },
  { field: "src", attr: "src", kind: "string" },
  { field: "imageaspect", attr: "imageaspect", kind: "string" },
  { field: "imagesize", attr: "imagesize", kind: "string" },
  { field: "imagealignshape", attr: "imagealignshape", kind: "trueFalse" },
  { field: "startarrow", attr: "startarrow", kind: "string" },
  { field: "startarrowwidth", attr: "startarrowwidth", kind: "string" },
  { field: "startarrowlength", attr: "startarrowlength", kind: "string" },
  { field: "endarrow", attr: "endarrow", kind: "string" },
  { field: "endarrowwidth", attr: "endarrowwidth", kind: "string" },
  { field: "endarrowlength", attr: "endarrowlength", kind: "string" },
  { field: "officeHref", attr: "o:href", kind: "string" },
  { field: "officeAltHref", attr: "o:althref", kind: "string" },
  { field: "officeTitle", attr: "o:title", kind: "string" },
  { field: "forcedash", attr: "o:forcedash", kind: "trueFalse" },
];

/** Serialize an o:left/top/right/bottom/column sub-stroke. */
function stringifyVmlStrokeChild(tag: string, opts: VmlStrokeChildOptions): string {
  return `<${tag}${stringifyVmlAttributes(opts as Record<string, unknown>, STROKE_CHILD_ATTRS)}/>`;
}

/** Parse an o: sub-stroke child of v:stroke. */
function parseVmlStrokeChild(el: XmlElement): VmlStrokeChildOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, STROKE_CHILD_ATTRS, out);
  return out as VmlStrokeChildOptions;
}

/** Serialize v:stroke. */
export function stringifyVmlStroke(opts: VmlStrokeOptions): string {
  const children: string[] = [];
  if (opts.leftStroke !== undefined)
    children.push(stringifyVmlStrokeChild("o:left", opts.leftStroke));
  if (opts.topStroke !== undefined) children.push(stringifyVmlStrokeChild("o:top", opts.topStroke));
  if (opts.rightStroke !== undefined)
    children.push(stringifyVmlStrokeChild("o:right", opts.rightStroke));
  if (opts.bottomStroke !== undefined)
    children.push(stringifyVmlStrokeChild("o:bottom", opts.bottomStroke));
  if (opts.columnStroke !== undefined)
    children.push(stringifyVmlStrokeChild("o:column", opts.columnStroke));
  const attrStr = stringifyVmlAttributes(opts as Record<string, unknown>, STROKE_ATTRS);
  return children.length > 0
    ? `<v:stroke${attrStr}>${children.join("")}</v:stroke>`
    : `<v:stroke${attrStr}/>`;
}

/** Parse a v:stroke element. */
export function parseVmlStroke(el: XmlElement): VmlStrokeOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, STROKE_ATTRS, out);
  for (const child of el.elements ?? []) {
    if (child.type !== "element") continue;
    switch (child.name) {
      case "o:left":
        out.leftStroke = parseVmlStrokeChild(child);
        break;
      case "o:top":
        out.topStroke = parseVmlStrokeChild(child);
        break;
      case "o:right":
        out.rightStroke = parseVmlStrokeChild(child);
        break;
      case "o:bottom":
        out.bottomStroke = parseVmlStrokeChild(child);
        break;
      case "o:column":
        out.columnStroke = parseVmlStrokeChild(child);
        break;
    }
  }
  return out as VmlStrokeOptions;
}
