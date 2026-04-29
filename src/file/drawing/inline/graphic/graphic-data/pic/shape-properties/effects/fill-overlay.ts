/**
 * Fill overlay effect for DrawingML shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_FillOverlayEffect
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import { createGradientFill } from "../fill/gradient-fill";
import type { GradientFillOptions } from "../fill/gradient-fill";
import { createGroupFill } from "../fill/group-fill";
import { createPatternFill } from "../fill/pattern-fill";
import type { PatternFillOptions } from "../fill/pattern-fill";
import { createNoFill } from "../outline/no-fill";
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
    /** Solid fill color */
    readonly solidFill?: SolidFillOptions;
    /** Gradient fill */
    readonly gradientFill?: GradientFillOptions;
    /** Pattern fill */
    readonly patternFill?: PatternFillOptions;
    /** Group fill (inherit from parent) */
    readonly groupFill?: boolean;
    /** No fill */
    readonly noFill?: boolean;
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
export const createFillOverlayEffect = (options: FillOverlayEffectOptions): XmlComponent => {
    let fillElement: XmlComponent;

    if (options.noFill) {
        fillElement = createNoFill();
    } else if (options.solidFill) {
        fillElement = createSolidFill(options.solidFill);
    } else if (options.gradientFill) {
        fillElement = createGradientFill(options.gradientFill);
    } else if (options.patternFill) {
        fillElement = createPatternFill(options.patternFill);
    } else if (options.groupFill) {
        fillElement = createGroupFill();
    } else {
        // Default to solid fill if nothing specified
        fillElement = createSolidFill({ value: "000000" });
    }

    return new BuilderElement<{ readonly blend: string }>({
        attributes: {
            blend: { key: "blend", value: options.blend },
        },
        children: [fillElement],
        name: "a:fillOverlay",
    });
};
