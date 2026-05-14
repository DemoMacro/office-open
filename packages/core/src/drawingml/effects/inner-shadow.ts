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
    readonly blurRadius?: number;
    /** Distance from shape edge in EMUs */
    readonly distance?: number;
    /** Direction angle in 60,000ths of a degree */
    readonly direction?: number;
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
        options.blurRadius !== undefined ||
        options.distance !== undefined ||
        options.direction !== undefined;

    if (!hasAttributes) {
        return new BuilderElement({
            children: [createColorElement(options.color)],
            name: "a:innerShdw",
        });
    }

    const attributePayload = {
        ...(options.blurRadius !== undefined && {
            blurRad: { key: "blurRad", value: options.blurRadius },
        }),
        ...(options.distance !== undefined && {
            dist: { key: "dist", value: options.distance },
        }),
        ...(options.direction !== undefined && {
            dir: { key: "dir", value: options.direction },
        }),
    };

    return new BuilderElement({
        attributes: attributePayload as never,
        children: [createColorElement(options.color)],
        name: "a:innerShdw",
    });
};
