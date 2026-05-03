/**
 * Inner shadow effect for DrawingML shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_InnerShadowEffect
 *
 * @module
 */
import { BuilderElement } from "../../xml-components";
import type { XmlComponent } from "../../xml-components";
import { createColorElement } from "../color/solid-fill";
import type { SolidFillOptions } from "../color/solid-fill";

/**
 * Options for inner shadow effect.
 */
export interface InnerShadowEffectOptions {
    /** Blur radius in EMUs */
    readonly blurRad?: number;
    /** Distance from shape edge in EMUs */
    readonly dist?: number;
    /** Direction angle in 60,000ths of a degree */
    readonly dir?: number;
    /** Shadow color */
    readonly color: SolidFillOptions;
}

/**
 * Creates an inner shadow effect element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_InnerShadowEffect">
 *   <xsd:sequence>
 *     <xsd:group ref="EG_ColorChoice" minOccurs="1" maxOccurs="1"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="blurRad" type="ST_PositiveCoordinate" default="0"/>
 *   <xsd:attribute name="dist" type="ST_PositiveCoordinate" default="0"/>
 *   <xsd:attribute name="dir" type="ST_PositiveFixedAngle" default="0"/>
 * </xsd:complexType>
 * ```
 */
export const createInnerShadowEffect = (options: InnerShadowEffectOptions): XmlComponent => {
    const hasAttributes =
        options.blurRad !== undefined || options.dist !== undefined || options.dir !== undefined;

    if (!hasAttributes) {
        return new BuilderElement({
            children: [createColorElement(options.color)],
            name: "a:innerShdw",
        });
    }

    const attributePayload = {
        ...(options.blurRad !== undefined && {
            blurRad: { key: "blurRad", value: options.blurRad },
        }),
        ...(options.dist !== undefined && {
            dist: { key: "dist", value: options.dist },
        }),
        ...(options.dir !== undefined && {
            dir: { key: "dir", value: options.dir },
        }),
    };

    return new BuilderElement({
        attributes: attributePayload as never,
        children: [createColorElement(options.color)],
        name: "a:innerShdw",
    });
};
