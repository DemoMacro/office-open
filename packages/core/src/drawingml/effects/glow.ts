/**
 * Glow effect for DrawingML shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_GlowEffect
 *
 * @module
 */
import { BuilderElement } from "../../xml-components";
import type { XmlComponent } from "../../xml-components";
import { createColorElement } from "../color/solid-fill";
import type { SolidFillOptions } from "../color/solid-fill";

/**
 * Options for glow effect.
 */
export interface GlowEffectOptions {
    /** Glow radius in EMUs */
    readonly rad?: number;
    /** Glow color */
    readonly color: SolidFillOptions;
}

/**
 * Creates a glow effect element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_GlowEffect">
 *   <xsd:sequence>
 *     <xsd:group ref="EG_ColorChoice" minOccurs="1" maxOccurs="1"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="rad" type="ST_PositiveCoordinate" use="optional" default="0"/>
 * </xsd:complexType>
 * ```
 */
export const createGlowEffect = (options: GlowEffectOptions): XmlComponent => {
    if (options.rad === undefined) {
        return new BuilderElement({
            children: [createColorElement(options.color)],
            name: "a:glow",
        });
    }

    return new BuilderElement<{ readonly rad: number }>({
        attributes: {
            rad: { key: "rad", value: options.rad },
        },
        children: [createColorElement(options.color)],
        name: "a:glow",
    });
};
