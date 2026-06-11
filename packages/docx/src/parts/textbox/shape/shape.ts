/**
 * VML shape module for WordprocessingML documents.
 *
 * Provides the VmlShapeStyle type and style-to-key mapping used by compile/
 * and parse paths. Runtime shape construction has been migrated to the
 * descriptor pipeline.
 *
 * References:
 * - https://c-rex.net/samples/ooxml/e1/Part3/OOXML_P3_Primer_OfficeArt_topic_ID0ELU5O.html
 * - http://webapp.docx4java.org/OnlineDemo/ecma376/VML/shape.html
 *
 * @module
 */
import type { LengthUnit } from "../types";

/**
 * Maps VmlShapeStyle property names to their corresponding CSS-style property names.
 * Used internally for converting TypeScript-friendly property names to VML style attributes.
 */
export const styleToKeyMap: Record<keyof VmlShapeStyle, string> = {
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
  zIndex: "z-index",
};

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
  /** Specifies the width of the containing block of the shape. Default is 0. */
  width: LengthUnit;
  /** Specifies the display order of overlapping shapes. Default is 0. */
  zIndex?: "auto" | number;
}
