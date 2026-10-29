/**
 * Fill overlay effect for DrawingML shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_FillOverlayEffect
 *
 * @module
 */
import { element } from "@office-open/xml";

import { NOOP_WRITE_CONTEXT, type WriteContext } from "../../descriptor";
import { xsdBlendMode } from "../../util/mappings";
import { fillDesc } from "../fill/fill-descriptors";
import type { FillOptions } from "../fill/fill-options";

/**
 * Blend modes for fill overlay effect.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_BlendMode">
 *   <xsd:restriction base="xsd:token">
 *     <xsd:enumeration value="over"/>
 *     <xsd:enumeration value="mult"/>
 *     <xsd:enumeration value="screen"/>
 *     <xsd:enumeration value="darken"/>
 *     <xsd:enumeration value="lighten"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 */
export const BlendMode = {
  /** Over blend mode */
  OVER: "over",
  /** Multiply blend mode */
  MULTIPLY: "multiply",
  /** Screen blend mode */
  SCREEN: "screen",
  /** Darken blend mode */
  DARKEN: "darken",
  /** Lighten blend mode */
  LIGHTEN: "lighten",
} as const;

/**
 * Options for fill overlay effect.
 */
export interface FillOverlayEffectOptions {
  /** Blend mode (required) */
  blend: (typeof BlendMode)[keyof typeof BlendMode];
  /** Overlay fill (CT_FillOverlayEffect EG_FillProperties child). */
  fill: FillOptions;
}

/**
 * Creates a fill overlay effect element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_FillOverlayEffect">
 *   <xsd:sequence>
 *     <xsd:group ref="EG_FillProperties" minOccurs="1" maxOccurs="1"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="blend" type="ST_BlendMode" use="required"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Solid fill overlay
 * const fillOverlay = createFillOverlayEffect({
 *   blend: BlendMode.MULTIPLY,
 *   solidFill: { value: "FF0000" },
 * });
 *
 * // Gradient fill overlay
 * const fillOverlay = createFillOverlayEffect({
 *   blend: BlendMode.SCREEN,
 *   gradientFill: { type: "linear", stops: [...] },
 * });
 * ```
 */
export const createFillOverlayEffect = (
  options: FillOverlayEffectOptions,
  context: WriteContext = NOOP_WRITE_CONTEXT,
): string =>
  element("a:fillOverlay", { blend: xsdBlendMode.to(options.blend) }, [
    fillDesc.stringify(options.fill, context) ?? "",
  ]);
