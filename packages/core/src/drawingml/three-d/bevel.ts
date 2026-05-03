/**
 * Bevel element for DrawingML 3D shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_Bevel
 *
 * @module
 */
import { BuilderElement } from "../../xml-components";

/**
 * Bevel preset types (12 variations).
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, ST_BevelPresetType
 */
export const BevelPresetType = {
    RELAXED_INSET: "relaxedInset",
    CIRCLE: "circle",
    SLOPE: "slope",
    CROSS: "cross",
    ANGLE: "angle",
    SOFT_ROUND: "softRound",
    CONVEX: "convex",
    COOL_SLANT: "coolSlant",
    DIVOT: "divot",
    RIBLET: "riblet",
    HARD_EDGE: "hardEdge",
    ART_DECO: "artDeco",
} as const;

/**
 * Options for a bevel element.
 */
export interface BevelOptions {
    /** Bevel width in EMUs (default 76200) */
    readonly w?: number;
    /** Bevel height in EMUs (default 76200) */
    readonly h?: number;
    /** Bevel preset type */
    readonly prst?: keyof typeof BevelPresetType;
}

/**
 * Creates a bevel element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Bevel">
 *   <xsd:attribute name="w" type="ST_PositiveCoordinate" default="76200"/>
 *   <xsd:attribute name="h" type="ST_PositiveCoordinate" default="76200"/>
 *   <xsd:attribute name="prst" type="ST_BevelPresetType" default="circle"/>
 * </xsd:complexType>
 * ```
 */
export const createBevel = (options?: BevelOptions) => {
    if (!options) {
        return new BuilderElement({ name: "a:bevelT" });
    }

    const attributes: Record<string, { readonly key: string; readonly value: string | number }> =
        {} as Record<string, { readonly key: string; readonly value: string | number }>;

    if (options.w !== undefined) {
        attributes.w = { key: "w", value: options.w };
    }
    if (options.h !== undefined) {
        attributes.h = { key: "h", value: options.h };
    }
    if (options.prst !== undefined) {
        attributes.prst = { key: "prst", value: BevelPresetType[options.prst] };
    }

    return new BuilderElement({
        attributes: attributes as never,
        name: "a:bevelT",
    });
};

/**
 * Creates a bottom bevel element (a:bevelB).
 */
export const createBottomBevel = (options?: BevelOptions) => {
    if (!options) {
        return new BuilderElement({ name: "a:bevelB" });
    }

    const attributes: Record<string, { readonly key: string; readonly value: string | number }> =
        {} as Record<string, { readonly key: string; readonly value: string | number }>;

    if (options.w !== undefined) {
        attributes.w = { key: "w", value: options.w };
    }
    if (options.h !== undefined) {
        attributes.h = { key: "h", value: options.h };
    }
    if (options.prst !== undefined) {
        attributes.prst = { key: "prst", value: BevelPresetType[options.prst] };
    }

    return new BuilderElement({
        attributes: attributes as never,
        name: "a:bevelB",
    });
};
