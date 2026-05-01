/**
 * Outer shadow effect for DrawingML shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_OuterShadowEffect
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import { createColorElement } from "../outline/solid-fill";
import type { SolidFillOptions } from "../outline/solid-fill";

/**
 * Rectangle alignment for shadow positioning.
 */
export const RectAlignment = {
    TOP_LEFT: "tl",
    TOP: "t",
    TOP_RIGHT: "tr",
    LEFT: "l",
    CENTER: "ctr",
    RIGHT: "r",
    BOTTOM_LEFT: "bl",
    BOTTOM: "b",
    BOTTOM_RIGHT: "br",
} as const;

/**
 * Options for outer shadow effect.
 */
export interface OuterShadowEffectOptions {
    /** Blur radius in EMUs */
    readonly blurRad?: number;
    /** Distance from shape in EMUs */
    readonly dist?: number;
    /** Direction angle in 60,000ths of a degree */
    readonly dir?: number;
    /** Horizontal scale percentage (e.g., 100000 = 100%) */
    readonly sx?: number;
    /** Vertical scale percentage */
    readonly sy?: number;
    /** Horizontal skew angle in 60,000ths of a degree */
    readonly kx?: number;
    /** Vertical skew angle */
    readonly ky?: number;
    /** Shadow alignment */
    readonly algn?: keyof typeof RectAlignment;
    /** Whether shadow rotates with shape */
    readonly rotWithShape?: boolean;
    /** Shadow color */
    readonly color: SolidFillOptions;
}

/**
 * Creates an outer shadow effect element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_OuterShadowEffect">
 *   <xsd:sequence>
 *     <xsd:group ref="EG_ColorChoice" minOccurs="1" maxOccurs="1"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="blurRad" type="ST_PositiveCoordinate" default="0"/>
 *   <xsd:attribute name="dist" type="ST_PositiveCoordinate" default="0"/>
 *   <xsd:attribute name="dir" type="ST_PositiveFixedAngle" default="0"/>
 *   <xsd:attribute name="sx" type="ST_Percentage" default="100%"/>
 *   <xsd:attribute name="sy" type="ST_Percentage" default="100%"/>
 *   <xsd:attribute name="kx" type="ST_FixedAngle" default="0"/>
 *   <xsd:attribute name="ky" type="ST_FixedAngle" default="0"/>
 *   <xsd:attribute name="algn" type="ST_RectAlignment" default="b"/>
 *   <xsd:attribute name="rotWithShape" type="xsd:boolean" default="true"/>
 * </xsd:complexType>
 * ```
 */
export const createOuterShadowEffect = (options: OuterShadowEffectOptions): XmlComponent => {
    const attributes: Record<string, { readonly key: string; readonly value: string | number }> =
        {};

    if (options.blurRad !== undefined) {
        attributes.blurRad = { key: "blurRad", value: options.blurRad };
    }
    if (options.dist !== undefined) {
        attributes.dist = { key: "dist", value: options.dist };
    }
    if (options.dir !== undefined) {
        attributes.dir = { key: "dir", value: options.dir };
    }
    if (options.sx !== undefined) {
        attributes.sx = { key: "sx", value: options.sx };
    }
    if (options.sy !== undefined) {
        attributes.sy = { key: "sy", value: options.sy };
    }
    if (options.kx !== undefined) {
        attributes.kx = { key: "kx", value: options.kx };
    }
    if (options.ky !== undefined) {
        attributes.ky = { key: "ky", value: options.ky };
    }
    if (options.algn !== undefined) {
        attributes.algn = { key: "algn", value: RectAlignment[options.algn] };
    }
    if (options.rotWithShape === false) {
        attributes.rotWithShape = { key: "rotWithShape", value: 0 };
    }

    const children: XmlComponent[] = [createColorElement(options.color)];

    return new BuilderElement({
        attributes: attributes as never,
        children,
        name: "a:outerShdw",
    });
};
