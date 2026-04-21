/**
 * Fill overlay effect for DrawingML shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_FillOverlayEffect
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import { createSolidFill } from "../outline/solid-fill";
import type { SolidFillOptions } from "../outline/solid-fill";

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
    MULTIPLY: "mult",
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
    readonly blend: (typeof BlendMode)[keyof typeof BlendMode];
    /** Fill color */
    readonly color: SolidFillOptions;
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
 * const fillOverlay = createFillOverlayEffect({
 *   blend: BlendMode.MULTIPLY,
 *   color: { value: "FF0000" },
 * });
 * ```
 */
export const createFillOverlayEffect = (options: FillOverlayEffectOptions): XmlComponent =>
    new BuilderElement<{ readonly blend: string }>({
        attributes: {
            blend: { key: "blend", value: options.blend },
        },
        children: [createSolidFill(options.color)],
        name: "a:fillOverlay",
    });
