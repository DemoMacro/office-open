/**
 * Preset shadow effect for DrawingML shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_PresetShadowEffect
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import { createColorElement } from "../outline/solid-fill";
import type { SolidFillOptions } from "../outline/solid-fill";

/**
 * Preset shadow types (20 variations).
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, ST_PresetShadowVal
 */
export const PresetShadowVal = {
    SHDW1: "shdw1",
    SHDW2: "shdw2",
    SHDW3: "shdw3",
    SHDW4: "shdw4",
    SHDW5: "shdw5",
    SHDW6: "shdw6",
    SHDW7: "shdw7",
    SHDW8: "shdw8",
    SHDW9: "shdw9",
    SHDW10: "shdw10",
    SHDW11: "shdw11",
    SHDW12: "shdw12",
    SHDW13: "shdw13",
    SHDW14: "shdw14",
    SHDW15: "shdw15",
    SHDW16: "shdw16",
    SHDW17: "shdw17",
    SHDW18: "shdw18",
    SHDW19: "shdw19",
    SHDW20: "shdw20",
} as const;

/**
 * Options for preset shadow effect.
 */
export interface PresetShadowEffectOptions {
    /** Preset shadow type (required) */
    readonly prst: keyof typeof PresetShadowVal;
    /** Distance from shape in EMUs */
    readonly dist?: number;
    /** Direction angle in 60,000ths of a degree */
    readonly dir?: number;
    /** Shadow color */
    readonly color: SolidFillOptions;
}

/**
 * Creates a preset shadow effect element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_PresetShadowEffect">
 *   <xsd:sequence>
 *     <xsd:group ref="EG_ColorChoice" minOccurs="1" maxOccurs="1"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="prst" type="ST_PresetShadowVal" use="required"/>
 *   <xsd:attribute name="dist" type="ST_PositiveCoordinate" default="0"/>
 *   <xsd:attribute name="dir" type="ST_PositiveFixedAngle" default="0"/>
 * </xsd:complexType>
 * ```
 */
export const createPresetShadowEffect = (options: PresetShadowEffectOptions): XmlComponent => {
    const attributes: Record<string, { readonly key: string; readonly value: string | number }> = {
        prst: { key: "prst", value: PresetShadowVal[options.prst] },
    };

    if (options.dist !== undefined) {
        attributes.dist = { key: "dist", value: options.dist };
    }
    if (options.dir !== undefined) {
        attributes.dir = { key: "dir", value: options.dir };
    }

    return new BuilderElement({
        attributes: attributes as never,
        children: [createColorElement(options.color)],
        name: "a:prstShdw",
    });
};
