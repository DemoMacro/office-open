/**
 * Reflection effect for DrawingML shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_ReflectionEffect
 *
 * @module
 */
import { BuilderElement } from "../../xml-components";

/**
 * Options for reflection effect.
 *
 * All properties are optional with XSD defaults.
 */
export interface ReflectionEffectOptions {
    /** Blur radius in EMUs */
    readonly blurRad?: number;
    /** Start opacity (fixed percentage, e.g., 100000 = 100%) */
    readonly stA?: number;
    /** Start position (fixed percentage) */
    readonly stPos?: number;
    /** End opacity (fixed percentage) */
    readonly endA?: number;
    /** End position (fixed percentage) */
    readonly endPos?: number;
    /** Distance from shape in EMUs */
    readonly dist?: number;
    /** Direction angle in 60,000ths of a degree */
    readonly dir?: number;
    /** Fade direction angle in 60,000ths of a degree */
    readonly fadeDir?: number;
    /** Horizontal scale percentage */
    readonly sx?: number;
    /** Vertical scale percentage */
    readonly sy?: number;
    /** Horizontal skew angle */
    readonly kx?: number;
    /** Vertical skew angle */
    readonly ky?: number;
    /** Alignment */
    readonly algn?: string;
    /** Whether reflection rotates with shape */
    readonly rotWithShape?: boolean;
}

/**
 * Creates a reflection effect element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_ReflectionEffect">
 *   <xsd:attribute name="blurRad" type="ST_PositiveCoordinate" default="0"/>
 *   <xsd:attribute name="stA" type="ST_PositiveFixedPercentage" default="100%"/>
 *   <xsd:attribute name="stPos" type="ST_PositiveFixedPercentage" default="0%"/>
 *   <xsd:attribute name="endA" type="ST_PositiveFixedPercentage" default="0%"/>
 *   <xsd:attribute name="endPos" type="ST_PositiveFixedPercentage" default="100%"/>
 *   <xsd:attribute name="dist" type="ST_PositiveCoordinate" default="0"/>
 *   <xsd:attribute name="dir" type="ST_PositiveFixedAngle" default="0"/>
 *   <xsd:attribute name="fadeDir" type="ST_PositiveFixedAngle" default="5400000"/>
 *   <xsd:attribute name="sx" type="ST_Percentage" default="100%"/>
 *   <xsd:attribute name="sy" type="ST_Percentage" default="100%"/>
 *   <xsd:attribute name="kx" type="ST_FixedAngle" default="0"/>
 *   <xsd:attribute name="ky" type="ST_FixedAngle" default="0"/>
 *   <xsd:attribute name="algn" type="ST_RectAlignment" default="b"/>
 *   <xsd:attribute name="rotWithShape" type="xsd:boolean" default="true"/>
 * </xsd:complexType>
 * ```
 */
export const createReflectionEffect = (options?: ReflectionEffectOptions) => {
    if (!options) {
        return new BuilderElement({ name: "a:reflection" });
    }

    const attributes: Record<string, { readonly key: string; readonly value: string | number }> =
        {};

    if (options.blurRad !== undefined) {
        attributes.blurRad = { key: "blurRad", value: options.blurRad };
    }
    if (options.stA !== undefined) {
        attributes.stA = { key: "stA", value: options.stA };
    }
    if (options.stPos !== undefined) {
        attributes.stPos = { key: "stPos", value: options.stPos };
    }
    if (options.endA !== undefined) {
        attributes.endA = { key: "endA", value: options.endA };
    }
    if (options.endPos !== undefined) {
        attributes.endPos = { key: "endPos", value: options.endPos };
    }
    if (options.dist !== undefined) {
        attributes.dist = { key: "dist", value: options.dist };
    }
    if (options.dir !== undefined) {
        attributes.dir = { key: "dir", value: options.dir };
    }
    if (options.fadeDir !== undefined) {
        attributes.fadeDir = { key: "fadeDir", value: options.fadeDir };
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
        attributes.algn = { key: "algn", value: options.algn };
    }
    if (options.rotWithShape === false) {
        attributes.rotWithShape = { key: "rotWithShape", value: 0 };
    }

    return new BuilderElement({
        attributes: attributes as never,
        name: "a:reflection",
    });
};
