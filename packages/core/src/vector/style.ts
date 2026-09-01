import type { Percentage, RelativeMeasure, UniversalMeasure } from "../util/values";
/**
 * VML style mini-language.
 *
 * VML elements (v:shape, v:rect, …) carry layout in a CSS-like `style`
 * attribute (`"position:absolute;width:100pt;height:50pt"`). This module owns
 * the shared vocabulary: the VmlShapeStyle type, the property-name mapping,
 * and the stringify/parse pair used by the shape element serializers in this
 * domain.
 *
 * References:
 * - https://c-rex.net/samples/ooxml/e1/Part3/OOXML_P3_Primer_OfficeArt_topic_ID0ELU5O.html
 * - http://webapp.docx4java.org/OnlineDemo/ecma376/VML/shape.html
 *
 * @module
 */
import { parseVmlTrueFalse } from "./attributes";

/**
 * Length value for VML style properties: "auto" (application-calculated),
 * number (typically points), percentage string, explicit-unit string
 * ("1cm"), or relative measure.
 */
export type LengthUnit = "auto" | number | Percentage | UniversalMeasure | RelativeMeasure;

/**
 * Maps VmlShapeStyle property names to their corresponding CSS-style property names.
 * Used internally for converting TypeScript-friendly property names to VML style attributes.
 */
const styleToKeyMap: Record<keyof VmlShapeStyle, string> = {
  flip: "flip",
  height: "height",
  left: "left",
  marginBottom: "margin-bottom",
  marginLeft: "margin-left",
  marginRight: "margin-right",
  marginTop: "margin-top",
  position: "position",
  positionHorizontal: "mso-position-horizontal",
  positionHorizontalRelative: "mso-position-horizontal-relative",
  positionVertical: "mso-position-vertical",
  positionVerticalRelative: "mso-position-vertical-relative",
  rotation: "rotation",
  top: "top",
  visibility: "visibility",
  width: "width",
  wrapDistanceBottom: "mso-wrap-distance-bottom",
  wrapDistanceLeft: "mso-wrap-distance-left",
  wrapDistanceRight: "mso-wrap-distance-right",
  wrapDistanceTop: "mso-wrap-distance-top",
  wrapEdited: "mso-wrap-edited",
  wrapStyle: "mso-wrap-style",
  fitShapeToText: "mso-fit-shape-to-text",
  directionAlt: "mso-direction-alt",
  zIndex: "z-index",
};

/**
 * Serialize a VmlShapeStyle object to the VML style attribute value
 * (`"width:100pt;height:50pt"`), translating property names via
 * {@link styleToKeyMap}.
 */
export function stringifyVmlStyle(style: VmlShapeStyle): string {
  return Object.entries(style)
    .map(([key, value]) => {
      // VML's CSS-style vocabulary spells booleans t/f (mso-fit-shape-to-text:t).
      const text = typeof value === "boolean" ? (value ? "t" : "f") : String(value);
      return `${styleToKeyMap[key as keyof VmlShapeStyle]}:${text}`;
    })
    .join(";");
}

/**
 * Parse a VML style attribute value into a CSS-name-keyed record
 * (`{ "mso-position-horizontal": "absolute", … }`). Keys stay in VML/CSS form;
 * callers read the entries they understand (width/height, wrap distances, …).
 */
export function parseVmlStyle(styleStr: string): Record<string, string> {
  const style: Record<string, string> = {};
  for (const part of styleStr.split(";")) {
    const [key, val] = part.split(":").map((s) => s.trim());
    if (key && val) style[key] = val;
  }
  return style;
}

/** Reverse of {@link styleToKeyMap} — CSS-style property name → VmlShapeStyle field. */
const keyToStyleMap = Object.fromEntries(
  Object.entries(styleToKeyMap).map(([field, key]) => [key, field]),
) as Record<string, keyof VmlShapeStyle>;

const NUMERIC = /^-?\d+(\.\d+)?$/;

/**
 * Parse a CSS-name-keyed style record into a VmlShapeStyle object (camelCase
 * fields, numeric and boolean literals coerced) — the inverse of
 * {@link stringifyVmlStyle}. Keys outside the map are dropped.
 */
export function parseVmlShapeStyle(record: Record<string, string>): VmlShapeStyle {
  const style: Partial<Record<keyof VmlShapeStyle, unknown>> = {};
  for (const [key, value] of Object.entries(record)) {
    const field = keyToStyleMap[key];
    if (!field) continue;
    if (field === "wrapEdited" || field === "fitShapeToText") {
      style[field] = parseVmlTrueFalse(value);
    } else if (NUMERIC.test(value)) {
      style[field] = Number(value);
    } else {
      style[field] = value;
    }
  }
  return style as unknown as VmlShapeStyle;
}

/**
 * VML shape styling properties for WordprocessingML documents.
 *
 * These properties map to CSS-style attributes on VML shape elements and control
 * the shape's appearance, layout, and interaction with surrounding text.
 */
export interface VmlShapeStyle {
  /** Specifies that the orientation of a shape is flipped. Default is no value. */
  flip?: "x" | "y" | "xy" | "yx";
  /** Specifies the height of the containing block of the shape. Default is 0. */
  height?: LengthUnit;
  /** Specifies the position of the left of the containing block relative to the element left of it. Default is 0. */
  left?: LengthUnit;
  /** Specifies the position of the bottom of the containing block relative to the shape anchor. Default is 0. */
  marginBottom?: LengthUnit;
  /** Specifies the position of the left of the containing block relative to the shape anchor. Default is 0. */
  marginLeft?: LengthUnit;
  /** Specifies the position of the right of the containing block relative to the shape anchor. Default is 0. */
  marginRight?: LengthUnit;
  /** Specifies the position of the top of the containing block relative to the shape anchor. Default is 0. */
  marginTop?: LengthUnit;
  /** Specifies the horizontal positioning data. Default is absolute. */
  positionHorizontal?: "absolute" | "left" | "center" | "right" | "inside" | "outside";
  /** Specifies relative horizontal position data. Default is text. */
  positionHorizontalRelative?: "margin" | "page" | "text" | "char";
  /** Specifies the vertical positioning data. Default is absolute. */
  positionVertical?: "absolute" | "left" | "center" | "right" | "inside" | "outside";
  /** Specifies relative vertical position data. Default is text. */
  positionVerticalRelative?: "margin" | "page" | "text" | "char";
  /** Specifies the distance from the bottom of the shape to the text that wraps around it. Default is 0 pt. */
  wrapDistanceBottom?: number;
  /** Specifies the distance from the left side of the shape to the text that wraps around it. Default is 0 pt. */
  wrapDistanceLeft?: number;
  /** Specifies the distance from the right side of the shape to the text that wraps around it. Default is 0 pt. */
  wrapDistanceRight?: number;
  /** Specifies the distance from the top of the shape to the text that wraps around it. Default is 0 pt. */
  wrapDistanceTop?: number;
  /** Specifies whether the wrap coordinates were customized by the user. Default is false. */
  wrapEdited?: boolean;
  /** Specifies that a shape's height adjusts to fit its text. Default is false. */
  fitShapeToText?: boolean;
  /** Textbox text-flow override — "auto" lets paragraph direction decide. */
  directionAlt?: "auto";
  /** Specifies the wrapping mode for text in shapes. Default is square. */
  wrapStyle?: "square" | "none";
  /** Specifies the type of positioning used to place an element. Default is static. */
  position?: "static" | "absolute" | "relative";
  /** Specifies the angle that a shape is rotated, in degrees. Default is 0. */
  rotation?: number;
  /** Specifies the position of the top of the containing block. Default is 0. */
  top?: LengthUnit;
  /** Specifies whether a shape is displayed. Default is inherit. */
  visibility?: "hidden" | "inherit";
  /** Specifies the width of the containing block. Default is 0. */
  width?: LengthUnit;
  /** Specifies the display order of overlapping shapes. Default is 0. */
  zIndex?: "auto" | number;
}
