/**
 * VML shape module for WordprocessingML documents.
 *
 * Provides functionality for creating VML shape elements with customizable styling and positioning.
 *
 * References:
 * - https://c-rex.net/samples/ooxml/e1/Part3/OOXML_P3_Primer_OfficeArt_topic_ID0ELU5O.html
 * - http://webapp.docx4java.org/OnlineDemo/ecma376/VML/shape.html
 *
 * @module
 */
import type { FileChild } from "@file/file-child";
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import type { LengthUnit } from "../types";
import { createVmlTextbox } from "../vml-textbox/vml-texbox";

const SHAPE_TYPE = "#_x0000_t202";

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
  zIndex: "z-index",
};

/**
 * Styling options for VML shapes.
 *
 * This type defines all available CSS-like styling properties for positioning, sizing,
 * and configuring VML shapes in WordprocessingML documents. These properties control
 * the shape's appearance, layout, and interaction with surrounding text.
 */
export interface VmlShapeStyle {
  /** Specifies that the orientation of a shape is flipped. Default is no value. */
  flip?: "x" | "y" | "xy" | "yx";
  /** Specifies the height of the containing block of the shape. Default is 0. It is specified in CSS units or, for elements in a group, in the coordinate system of the parent element. */
  height?: LengthUnit;
  /** Specifies the position of the left of the containing block of the shape relative to the element left of it in the flow of the page. Default is 0. It is specified in CSS units or, for elements in a group, in the coordinate system of the parent element. This property shall not be used for shapes anchored inline. */
  left?: LengthUnit;
  /** Specifies the position of the bottom of the containing block of the shape relative to the shape anchor. Default is 0. It is specified in CSS units or, for elements in a group, in the coordinate system of the parent element. */
  marginBottom?: LengthUnit;
  /** Specifies the position of the left of the containing block of the shape relative to the shape anchor. Default is 0. It is specified in CSS units or, for elements in a group, in the coordinate system of the parent element. */
  marginLeft?: LengthUnit;
  /** Specifies the position of the right of the containing block of the shape relative to the shape anchor. Default is 0. It is specified in CSS units or, for elements in a group, in the coordinate system of the parent element. */
  marginRight?: LengthUnit;
  /** Specifies the position of the top of the containing block of the shape relative to the shape anchor. Default is 0. It is specified in CSS units or, for elements in a group, in the coordinate system of the parent element. */
  marginTop?: LengthUnit;
  /** Specifies the horizontal positioning data for objects in WordprocessingML documents. Default is absolute. */
  positionHorizontal?: "absolute" | "left" | "center" | "right" | "inside" | "outside";
  /** Specifies relative horizontal position data for objects in WordprocessingML documents. This modifies the mso-position-horizontal property. Default is text. */
  positionHorizontalRelative?: "margin" | "page" | "text" | "char";
  /** Specifies the vertical positioning data for objects in WordprocessingML documents. Default is absolute. */
  positionVertical?: "absolute" | "left" | "center" | "right" | "inside" | "outside";
  /** Specifies relative vertical position data for objects in WordprocessingML documents. This modifies the mso-position-vertical property. Default is text. */
  positionVerticalRelative?: "margin" | "page" | "text" | "char";
  /** Specifies the distance from the bottom of the shape to the text that wraps around it. Default is 0 pt. Note that this property is different from the CSS margin property, which changes the origin of the shape to include the margin areas. This property does not change the origin. */
  wrapDistanceBottom?: number;
  /** Specifies the distance from the left side of the shape to the text that wraps around it. Default is 0 pt. Note that this property is different from the CSS margin property, which changes the origin of the shape to include the margin areas. This property does not change the origin. */
  wrapDistanceLeft?: number;
  /** Specifies the distance from the right side of the shape to the text that wraps around it. Default is 0 pt. Note that this property is different from the CSS margin property, which changes the origin of the shape to include the margin areas. This property does not change the origin. */
  wrapDistanceRight?: number;
  /** Specifies the distance from the top of the shape to the text that wraps around it. Default is 0 pt. Note that this property is different from the CSS margin property, which changes the origin of the shape to include the margin areas. This property does not change the origin. */
  wrapDistanceTop?: number;
  /** Specifies whether the wrap coordinates were customized by the user. If the wrap coordinates are generated by an editor, this property is true; otherwise they were customized by a user. Default is false. */
  wrapEdited?: boolean;
  /** Specifies the wrapping mode for text in shapes in WordprocessingML documents. Default is square. */
  wrapStyle?: "square" | "none";
  /** Specifies the type of positioning used to place an element. Default is static. When the element is contained inside a group, this property must be absolute. */
  position?: "static" | "absolute" | "relative";
  /** Specifies the angle that a shape is rotated, in degrees. Default is 0. Positive angles are clockwise. */
  rotation?: number;
  /** Specifies the position of the top of the containing block of the shape relative to the element above it in the flow of the page. Default is 0. It is specified in CSS units or, for elements in a group, in the coordinate system of the parent element. This property shall not be used for shapes anchored inline. */
  top?: LengthUnit;
  /** Specifies whether a shape is displayed. Only inherit and hidden are used; any other values are mapped to inherit. Default is inherit. */
  visibility?: "hidden" | "inherit";
  /** Specifies the width of the containing block of the shape. Default is 0. It is specified in CSS units or, for elements in a group, in the coordinate system of the parent element. */
  width: LengthUnit;
  /** Specifies the display order of overlapping shapes. Default is 0. This property shall not be used for shapes anchored inline. */
  zIndex?: "auto" | number;
}

/**
 * Formats VmlShapeStyle object into a CSS-style string for VML shape attributes.
 *
 * @param style - The VmlShapeStyle object to format
 * @returns A CSS-style string (e.g., "width:100pt;height:50pt;") or undefined if no style provided
 * @internal
 */
const formatShapeStyle = (style?: VmlShapeStyle): string | undefined =>
  style
    ? Object.entries(style)
        .map(([key, value]) => `${styleToKeyMap[key as keyof VmlShapeStyle]}:${value}`)
        .join(";")
    : undefined;

/**
 * Options for creating a VML shape.
 *
 * @property id - Unique identifier for the shape
 * @property children - Array of paragraph children to include in the shape's textbox
 * @property type - VML shape type identifier (default: "#_x0000_t202" for text rectangle)
 * @property style - Styling properties for the shape
 */
interface ShapeOptions {
  /** Unique identifier for the shape */
  id: string;
  /** Array of block-level children to include in the shape's textbox */
  children?: FileChild[];
  /** VML shape type identifier (default: "#_x0000_t202" for text rectangle) */
  type?: string;
  /** Styling properties for the shape */
  style?: VmlShapeStyle;
  /** Fill color for the shape (VML: fillcolor) */
  fillColor?: string;
  /** Stroke color for the shape (VML: strokecolor) */
  strokeColor?: string;
  /** Stroke weight for the shape (VML: strokeweight) */
  strokeWeight?: string;
  /** Whether the shape is filled (VML: filled) */
  filled?: boolean;
  /** Whether the shape is stroked (VML: stroked) */
  stroked?: boolean;
  /** Coordinate origin, e.g. "0,0" (VML: coordorigin) */
  coordOrigin?: string;
  /** Coordinate size, e.g. "21600,21600" (VML: coordsize) */
  coordSize?: string;
  /** Inset pen mode (VML: insetpen) */
  insetPen?: boolean;
  /** Adjustment values for the shape (VML: adj) */
  adjustment?: string;
  /** Path definition for the shape (VML: path) */
  path?: string;
  /** Arc size for rounded rectangles, e.g. "10923f" (VML: arcsize) */
  arcSize?: string;
  /** Whether the shape is printed (VML: print) */
  print?: boolean;
  /** Equation XML content for the shape (VML: equationxml) */
  equationXml?: string;
}

/**
 * Creates a VML shape element with textbox content.
 *
 * The VML shape element (v:shape) represents a vector graphics shape in WordprocessingML documents.
 * This function creates shapes configured for text display (textbox shapes), which are commonly
 * used for creating floating text boxes with custom positioning and styling.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Shape">
 *   <xsd:choice maxOccurs="unbounded">
 *     <xsd:group ref="EG_ShapeElements"/>
 *     <xsd:element ref="o:ink"/>
 *     <xsd:element ref="pvml:iscomment"/>
 *     <xsd:element ref="o:equationxml"/>
 *   </xsd:choice>
 *   <xsd:attributeGroup ref="AG_AllCoreAttributes"/>
 *   <xsd:attributeGroup ref="AG_AllShapeAttributes"/>
 *   <xsd:attributeGroup ref="AG_Type"/>
 *   <xsd:attributeGroup ref="AG_Adj"/>
 *   <xsd:attributeGroup ref="AG_Path"/>
 *   <xsd:attribute ref="o:gfxdata"/>
 *   <xsd:attribute name="equationxml" type="xsd:string" use="optional"/>
 * </xsd:complexType>
 * ```
 *
 * @param options - Configuration options for the shape
 * @returns An XmlComponent representing the v:shape element
 *
 * @example
 * ```typescript
 * const shape = createShape({
 *   id: "textbox1",
 *   children: [new TextRun("Hello World")],
 *   style: {
 *     width: "3in",
 *     height: "1in",
 *     position: "absolute",
 *     left: "1in",
 *     top: "1in"
 *   }
 * });
 * ```
 */
export const createShape = ({
  id,
  children,
  type = SHAPE_TYPE,
  style,
  fillColor,
  strokeColor,
  strokeWeight,
  filled,
  stroked,
  coordOrigin,
  coordSize,
  insetPen,
  adjustment,
  path,
  arcSize,
  print,
  equationXml,
}: ShapeOptions): XmlComponent =>
  new BuilderElement({
    name: "v:shape",
    attributes: {
      id: {
        key: "id",
        value: id,
      },
      style: {
        key: "style",
        value: formatShapeStyle(style),
      },
      type: {
        key: "type",
        value: type,
      },
      fillColor: {
        key: "fillcolor",
        value: fillColor,
      },
      strokeColor: {
        key: "strokecolor",
        value: strokeColor,
      },
      strokeWeight: {
        key: "strokeweight",
        value: strokeWeight,
      },
      filled: {
        key: "filled",
        value: filled !== undefined ? (filled ? "true" : "false") : undefined,
      },
      stroked: {
        key: "stroked",
        value: stroked !== undefined ? (stroked ? "true" : "false") : undefined,
      },
      coordOrigin: {
        key: "coordorigin",
        value: coordOrigin,
      },
      coordSize: {
        key: "coordsize",
        value: coordSize,
      },
      insetPen: {
        key: "insetpen",
        value: insetPen !== undefined ? (insetPen ? "true" : "false") : undefined,
      },
      adjustment: {
        key: "adj",
        value: adjustment,
      },
      path: {
        key: "path",
        value: path,
      },
      arcSize: {
        key: "arcsize",
        value: arcSize,
      },
      print: {
        key: "print",
        value: print !== undefined ? (print ? "true" : "false") : undefined,
      },
      equationXml: {
        key: "equationxml",
        value: equationXml,
      },
    },
    children: [createVmlTextbox({ children, style: "mso-fit-shape-to-text:t;" })],
  });
