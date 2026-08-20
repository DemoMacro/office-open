import { element } from "@office-open/xml";

import { convertToEmu } from "../../util/converters";
import { xsdCompoundLine, xsdLineCap, xsdPenAlignment } from "../../util/mappings";
import { stripColorHashPrefix } from "../../util/values";
/**
 * Outline (line) properties for DrawingML shapes.
 *
 * This module provides support for configuring outline properties including
 * width, cap style, compound line types, fill properties, dash, and join.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_LineProperties
 *
 * @module
 */
import type { UniversalMeasure } from "../../util/values";
import { createSolidFill } from "../color/solid-fill";
import type { SolidFillOptions } from "../color/solid-fill";
import { createGradientFill } from "../fill/gradient-fill";
import type { GradientFillOptions } from "../fill/gradient-fill";
import { createNoFill } from "../fill/no-fill";
import { createPatternFill } from "../fill/pattern-fill";
import type { PatternFillOptions } from "../fill/pattern-fill";
import { createCustomDash } from "./custom-dash";
import type { DashStop } from "./custom-dash";
import { createLineEnd } from "./line-end";
import type { LineEndOptions } from "./line-end";

// <xsd:complexType name="CT_LineProperties">
//     <xsd:sequence>
//         <xsd:group ref="EG_LineFillProperties" minOccurs="0"/>
//         <xsd:group ref="EG_LineDashProperties" minOccurs="0"/>
//         <xsd:group ref="EG_LineJoinProperties" minOccurs="0"/>
//     </xsd:sequence>
//     <xsd:attribute name="w" use="optional" type="a:ST_LineWidth"/>
//     <xsd:attribute name="cap" use="optional" type="ST_LineCap"/>
//     <xsd:attribute name="cmpd" use="optional" type="ST_CompoundLine"/>
//     <xsd:attribute name="algn" use="optional" type="ST_PenAlignment"/>
// </xsd:complexType>

// <xsd:simpleType name="ST_LineCap">
//     <xsd:restriction base="xsd:string">
//     <xsd:enumeration value="rnd"/>
//     <xsd:enumeration value="sq"/>
//     <xsd:enumeration value="flat"/>
//     </xsd:restriction>
// </xsd:simpleType>

/**
 * Line cap styles for outline endpoints.
 *
 * Defines how the ends of a line are rendered.
 */
export type LineCap = "round" | "square" | "flat";

// <xsd:simpleType name="ST_CompoundLine">
//     <xsd:restriction base="xsd:string">
//     <xsd:enumeration value="sng"/>
//     <xsd:enumeration value="dbl"/>
//     <xsd:enumeration value="thickThin"/>
//     <xsd:enumeration value="thinThick"/>
//     <xsd:enumeration value="tri"/>
//     </xsd:restriction>
// </xsd:simpleType>

/**
 * Compound line types for outlines.
 *
 * Defines the structure of compound lines (single, double, etc.).
 */
export type CompoundLine = "single" | "double" | "thickThin" | "thinThick" | "triple";

// <xsd:simpleType name="ST_PenAlignment">
//     <xsd:restriction base="xsd:string">
//     <xsd:enumeration value="ctr"/>
//     <xsd:enumeration value="in"/>
//     </xsd:restriction>
// </xsd:simpleType>

/**
 * Pen alignment options for outline positioning.
 *
 * Defines how the outline is aligned relative to the shape edge.
 */
export type PenAlignment = "center" | "inside";

/**
 * Preset dash styles for outlines.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_PresetLineDashVal">
 *   <xsd:restriction base="xsd:token">
 *     <xsd:enumeration value="solid"/>
 *     <xsd:enumeration value="dot"/>
 *     <xsd:enumeration value="dash"/>
 *     <xsd:enumeration value="lgDash"/>
 *     <xsd:enumeration value="dashDot"/>
 *     <xsd:enumeration value="lgDashDot"/>
 *     <xsd:enumeration value="lgDashDotDot"/>
 *     <xsd:enumeration value="sysDash"/>
 *     <xsd:enumeration value="sysDot"/>
 *     <xsd:enumeration value="sysDashDot"/>
 *     <xsd:enumeration value="sysDashDotDot"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 */
export type PresetDash =
  | "solid"
  | "dot"
  | "dash"
  | "lgDash"
  | "dashDot"
  | "lgDashDot"
  | "lgDashDotDot"
  | "sysDash"
  | "sysDot"
  | "sysDashDot"
  | "sysDashDotDot";

/**
 * Line join styles.
 */
export type LineJoin = "round" | "bevel" | "miter";

/**
 * Attributes for configuring outline properties.
 */
export interface OutlineProperties {
  /** Line width in EMUs (English Metric Units) or universal measure (e.g., "1pt", "2mm") */
  width?: number | UniversalMeasure;
  /** Line cap style */
  cap?: LineCap;
  /** Compound line type */
  compoundLine?: CompoundLine;
  /** Pen alignment */
  alignment?: PenAlignment;
  /**
   * Preset dash style.
   *
   * Mutually exclusive with `customDash` — only one can be specified.
   */
  dash?: PresetDash;
  /**
   * Custom dash pattern (list of dash/space stops).
   *
   * Mutually exclusive with `dash` — only one can be specified.
   */
  customDash?: readonly DashStop[];
  /** Line join style */
  join?: LineJoin;
  /** Miter limit (only when join is MITER) */
  miterLimit?: number;
  /** Line start arrow/head */
  headEnd?: LineEndOptions;
  /** Line end arrow/tail */
  tailEnd?: LineEndOptions;
  /** Trailing a:extLst inner XML for line-property extensions. */
  ext?: string;
}

/**
 * Fill properties for outline (EG_LineFillProperties).
 *
 * Supports noFill, solidFill, gradFill, and pattFill per XSD.
 * blipFill and grpFill are not applicable to lines.
 */
export interface OutlineFillProperties {
  /** Fill type */
  type?: "noFill" | "solidFill" | "gradFill" | "pattFill";
  /**
   * Color definition (required when type is "solidFill"). A bare string is an
   * sRGB hex convenience sugar (e.g. `"FF0000"`); stringify coerces it to
   * `{ value }` and infers `type: "solidFill"` when `type` is omitted. Parse
   * always emits the normalized `{ value }` form.
   */
  color?: SolidFillOptions | string;
  /** Gradient fill options (required when type is "gradFill") */
  gradientFill?: GradientFillOptions;
  /** Pattern fill options (required when type is "pattFill") */
  patternFill?: PatternFillOptions;
}

/**
 * Complete outline configuration options.
 *
 * Combines outline attributes with fill properties.
 */
export type OutlineOptions = OutlineProperties & OutlineFillProperties;

/**
 * Creates the fill child element for an outline.
 *
 * Returns null when no fill type is specified (OOXML allows outline without fill).
 */
const createOutlineFill = (options: OutlineOptions): string | null => {
  if (options.type === "noFill") {
    return createNoFill();
  }
  // Bare-string color is an sRGB hex sugar; coerce and infer solidFill.
  const resolvedColor =
    typeof options.color === "string"
      ? { value: stripColorHashPrefix(options.color) }
      : options.color;
  const fillType = options.type ?? (resolvedColor ? "solidFill" : undefined);
  if (fillType === "solidFill" && resolvedColor) {
    return createSolidFill(resolvedColor);
  }
  if (fillType === "gradFill" && options.gradientFill) {
    return createGradientFill(options.gradientFill);
  }
  if (fillType === "pattFill" && options.patternFill) {
    return createPatternFill(options.patternFill);
  }
  return null;
};

/**
 * Creates an outline element for DrawingML shapes.
 *
 * The outline element specifies the line properties for the shape border,
 * including width, cap style, compound line type, alignment, dash, join, and fill.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_LineProperties">
 *   <xsd:sequence>
 *     <xsd:group ref="EG_FillProperties" minOccurs="0"/>
 *     <xsd:group ref="EG_LineDashProperties" minOccurs="0"/>
 *     <xsd:group ref="EG_LineJoinProperties" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="w" use="optional" type="a:ST_LineWidth"/>
 *   <xsd:attribute name="cap" use="optional" type="ST_LineCap"/>
 *   <xsd:attribute name="cmpd" use="optional" type="ST_CompoundLine"/>
 *   <xsd:attribute name="algn" use="optional" type="ST_PenAlignment"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Outline with RGB color and dash
 * const outline = createOutline({
 *   width: 9525,
 *   type: "solidFill",
 *   color: { value: "FF0000" },
 *   dash: "dash",
 * });
 * ```
 */
export const createOutline = (options: OutlineOptions): string => {
  const children: string[] = [];

  // Fill (optional per OOXML spec)
  const fill = createOutlineFill(options);
  if (fill) {
    children.push(fill);
  }

  // Dash (prstDash and custDash are mutually exclusive per XSD choice)
  if (options.customDash !== undefined) {
    children.push(createCustomDash(options.customDash));
  } else if (options.dash !== undefined) {
    children.push(`<a:prstDash val="${options.dash}"/>`);
  }

  // Join
  if (options.join !== undefined) {
    if (options.join === "miter" && options.miterLimit !== undefined) {
      children.push(`<a:miter lim="${options.miterLimit}"/>`);
    } else {
      children.push(`<a:${options.join}/>`);
    }
  }

  // Line end markers (arrows)
  if (options.headEnd) {
    children.push(createLineEnd("a:headEnd", options.headEnd));
  }
  if (options.tailEnd) {
    children.push(createLineEnd("a:tailEnd", options.tailEnd));
  }

  return element(
    "a:ln",
    {
      algn: options.alignment ? xsdPenAlignment.to(options.alignment) : undefined,
      cap: options.cap ? xsdLineCap.to(options.cap) : undefined,
      cmpd: options.compoundLine ? xsdCompoundLine.to(options.compoundLine) : undefined,
      w: options.width !== undefined ? convertToEmu(options.width) : undefined,
    },
    children,
  );
};
