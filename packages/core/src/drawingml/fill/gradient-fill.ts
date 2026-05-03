/**
 * Gradient fill element for DrawingML shapes.
 *
 * This module provides gradient fill support with linear and path shading.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_GradientFillProperties
 *
 * @module
 */
import { BuilderElement } from "../../xml-components";
import type { XmlComponent } from "../../xml-components";
import type { SolidFillOptions } from "../color/solid-fill";
import { createColorElement } from "../color/solid-fill";

/**
 * Gradient stop position (0-100000, representing 0%-100%).
 */
export interface IGradientStop {
    /** Position of the color stop (0-100000) */
    readonly position: number;
    /** Color at this stop */
    readonly color: SolidFillOptions;
}

/**
 * Path shade type for radial gradients.
 */
export const PathShadeType = {
    /** Follow shape path */
    SHAPE: "shape",
    /** Circular gradient */
    CIRCLE: "circle",
    /** Rectangular gradient */
    RECT: "rect",
} as const;

/**
 * Tile flip mode for gradient fill.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_TileFlipMode">
 *   <xsd:restriction base="xsd:string">
 *     <xsd:enumeration value="none"/>
 *     <xsd:enumeration value="x"/>
 *     <xsd:enumeration value="y"/>
 *     <xsd:enumeration value="xy"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 */
export const TileFlipMode = {
    /** No flip */
    NONE: "none",
    /** Flip horizontally */
    X: "x",
    /** Flip vertically */
    Y: "y",
    /** Flip both horizontally and vertically */
    XY: "xy",
} as const;

/**
 * Options for linear gradient shading.
 */
export interface LinearShadeOptions {
    /** Angle in 60,000ths of a degree (e.g., 5400000 = 90°) */
    readonly angle?: number;
    /** Whether the angle scales with the shape */
    readonly scaled?: boolean;
}

/**
 * Relative rectangle (CT_RelativeRect) with percentage offsets.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_RelativeRect">
 *   <xsd:attribute name="l" type="ST_Percentage" default="0%"/>
 *   <xsd:attribute name="t" type="ST_Percentage" default="0%"/>
 *   <xsd:attribute name="r" type="ST_Percentage" default="0%"/>
 *   <xsd:attribute name="b" type="ST_Percentage" default="0%"/>
 * </xsd:complexType>
 * ```
 */
export interface RelativeRect {
    /** Left offset percentage (e.g., "0%") */
    readonly left?: string;
    /** Top offset percentage (e.g., "0%") */
    readonly top?: string;
    /** Right offset percentage (e.g., "0%") */
    readonly right?: string;
    /** Bottom offset percentage (e.g., "0%") */
    readonly bottom?: string;
}

/**
 * Options for path (radial) gradient shading.
 */
export interface PathShadeOptions {
    /** Path type */
    readonly path?: keyof typeof PathShadeType;
    /**
     * Fill-to rectangle for path gradient.
     *
     * Defines the rectangle to which the gradient fills.
     */
    readonly fillToRect?: RelativeRect;
}

/**
 * Gradient shade options (linear or path).
 */
export type GradientShadeOptions = LinearShadeOptions | PathShadeOptions;

/**
 * Options for gradient fill.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_GradientFillProperties">
 *   <xsd:sequence>
 *     <xsd:element name="gsLst" type="CT_GradientStopList" minOccurs="0"/>
 *     <xsd:group ref="EG_ShadeProperties" minOccurs="0"/>
 *     <xsd:element name="tileRect" type="CT_RelativeRect" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="flip" type="ST_TileFlipMode" use="optional"/>
 *   <xsd:attribute name="rotWithShape" type="xsd:boolean" use="optional"/>
 * </xsd:complexType>
 * ```
 */
export interface GradientFillOptions {
    /** Gradient color stops (minimum 2) */
    readonly stops: readonly IGradientStop[];
    /** Shade type (linear or path) */
    readonly shade?: GradientShadeOptions;
    /**
     * Tile flip mode.
     *
     * Controls how the gradient is flipped when tiled.
     */
    readonly flip?: keyof typeof TileFlipMode;
    /**
     * Tile rectangle for gradient tiling.
     *
     * Defines the rectangle used for gradient tiling.
     */
    readonly tileRect?: RelativeRect;
    /** Whether gradient rotates with the shape */
    readonly rotateWithShape?: boolean;
}

/**
 * Creates a gradient stop element (a:gs).
 *
 * @example
 * ```typescript
 * createGradientStop({ position: 0, color: { value: "FF0000" } });
 * createGradientStop({ position: 100000, color: { value: "0000FF" } });
 * ```
 */
export const createGradientStop = (stop: IGradientStop): XmlComponent =>
    new BuilderElement<{ readonly pos: number }>({
        attributes: {
            pos: { key: "pos", value: stop.position },
        },
        children: [createColorElement(stop.color)],
        name: "a:gs",
    });

/**
 * Creates a relative rect element.
 */
const createRelativeRect = (name: string, rect?: RelativeRect): XmlComponent =>
    new BuilderElement<{
        readonly l?: string;
        readonly t?: string;
        readonly r?: string;
        readonly b?: string;
    }>({
        attributes: {
            l: { key: "l", value: rect?.left },
            t: { key: "t", value: rect?.top },
            r: { key: "r", value: rect?.right },
            b: { key: "b", value: rect?.bottom },
        },
        name,
    });

/**
 * Creates the shade element (a:lin or a:path).
 */
const createShadeElement = (shade: GradientShadeOptions): XmlComponent => {
    if ("angle" in shade) {
        return new BuilderElement<{ readonly ang?: number; readonly scaled?: boolean }>({
            attributes: {
                ang: { key: "ang", value: shade.angle },
                scaled: { key: "scaled", value: shade.scaled },
            },
            name: "a:lin",
        });
    }
    const pathShade = shade as PathShadeOptions;
    const children: XmlComponent[] = [];

    if (pathShade.fillToRect) {
        children.push(createRelativeRect("a:fillToRect", pathShade.fillToRect));
    }

    return new BuilderElement<{ readonly path?: string }>({
        attributes: {
            path: {
                key: "path",
                value: pathShade.path ? PathShadeType[pathShade.path] : undefined,
            },
        },
        children,
        name: "a:path",
    });
};

/**
 * Creates a gradient fill element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_GradientFillProperties">
 *   <xsd:sequence>
 *     <xsd:element name="gsLst" type="CT_GradientStopList" minOccurs="0"/>
 *     <xsd:group ref="EG_ShadeProperties" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="rotWithShape" type="xsd:boolean" use="optional"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Linear gradient from red to blue
 * createGradientFill({
 *   stops: [
 *     { position: 0, color: { value: "FF0000" } },
 *     { position: 100000, color: { value: "0000FF" } },
 *   ],
 *   shade: { angle: 5400000 },
 * });
 * ```
 */
export const createGradientFill = (options: GradientFillOptions): XmlComponent => {
    const children: XmlComponent[] = [];

    // Gradient stop list
    children.push(
        new BuilderElement({
            children: options.stops.map(createGradientStop),
            name: "a:gsLst",
        }),
    );

    // Shade properties (a:lin or a:path)
    if (options.shade) {
        children.push(createShadeElement(options.shade));
    }

    // Tile rectangle
    if (options.tileRect) {
        children.push(createRelativeRect("a:tileRect", options.tileRect));
    }

    return new BuilderElement<{
        readonly flip?: string;
        readonly rotWithShape?: boolean;
    }>({
        attributes: {
            flip: {
                key: "flip",
                value: options.flip ? TileFlipMode[options.flip] : undefined,
            },
            rotWithShape: { key: "rotWithShape", value: options.rotateWithShape },
        },
        children,
        name: "a:gradFill",
    });
};
