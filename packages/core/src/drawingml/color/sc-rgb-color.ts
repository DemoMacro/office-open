/**
 * ScRGB color element for DrawingML shapes.
 *
 * This module provides scRGB color support using percentage-based RGB values.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_ScRgbColor
 *
 * @module
 */
import { BuilderElement } from "../../xml-components";
import type { XmlComponent } from "../../xml-components";
import type { ColorTransformOptions } from "./color-transform";
import { createColorTransforms } from "./color-transform";

/**
 * Options for scRGB color.
 */
export interface ScRgbColorOptions {
    /** Red percentage (e.g., "50%" or "100%") */
    readonly r: string;
    /** Green percentage (e.g., "50%" or "100%") */
    readonly g: string;
    /** Blue percentage (e.g., "50%" or "100%") */
    readonly b: string;
    /** Optional color transforms */
    readonly transforms?: ColorTransformOptions;
}

/**
 * Creates an scRGB color element.
 *
 * Specifies a color using percentage-based RGB values.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_ScRgbColor">
 *   <xsd:sequence>
 *     <xsd:group ref="EG_ColorTransform" minOccurs="0" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="r" type="ST_Percentage" use="required"/>
 *   <xsd:attribute name="g" type="ST_Percentage" use="required"/>
 *   <xsd:attribute name="b" type="ST_Percentage" use="required"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * const redColor = createScRgbColor({ r: "100%", g: "0%", b: "0%" });
 * // With alpha transform
 * const semiRed = createScRgbColor({ r: "100%", g: "0%", b: "0%", transforms: { alpha: 50000 } });
 * ```
 */
export const createScRgbColor = (options: ScRgbColorOptions): XmlComponent => {
    const transforms = options.transforms ? createColorTransforms(options.transforms) : [];
    return new BuilderElement<ScRgbColorOptions>({
        attributes: {
            r: {
                key: "r",
                value: options.r,
            },
            g: {
                key: "g",
                value: options.g,
            },
            b: {
                key: "b",
                value: options.b,
            },
        },
        children: [...transforms],
        name: "a:scrgbClr",
    });
};
