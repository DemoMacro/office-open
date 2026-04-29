/**
 * Line end (arrow) properties for DrawingML outlines.
 *
 * This module provides support for line end markers (arrows) on shape outlines.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_LineEndProperties
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

/**
 * Line end types (arrow head styles).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_LineEndType">
 *   <xsd:restriction base="xsd:token">
 *     <xsd:enumeration value="none"/>
 *     <xsd:enumeration value="triangle"/>
 *     <xsd:enumeration value="stealth"/>
 *     <xsd:enumeration value="diamond"/>
 *     <xsd:enumeration value="oval"/>
 *     <xsd:enumeration value="arrow"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @publicApi
 */
export const LineEndType = {
    /** No line end */
    NONE: "none",
    /** Triangle arrow */
    TRIANGLE: "triangle",
    /** Stealth arrow (filled triangle) */
    STEALTH: "stealth",
    /** Diamond shape */
    DIAMOND: "diamond",
    /** Oval shape */
    OVAL: "oval",
    /** Simple arrow */
    ARROW: "arrow",
} as const;

/**
 * Line end width options.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_LineEndWidth">
 *   <xsd:restriction base="xsd:token">
 *     <xsd:enumeration value="sm"/>
 *     <xsd:enumeration value="med"/>
 *     <xsd:enumeration value="lg"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @publicApi
 */
export const LineEndWidth = {
    /** Small width */
    SMALL: "sm",
    /** Medium width */
    MEDIUM: "med",
    /** Large width */
    LARGE: "lg",
} as const;

/**
 * Line end length options.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:simpleType name="ST_LineEndLength">
 *   <xsd:restriction base="xsd:token">
 *     <xsd:enumeration value="sm"/>
 *     <xsd:enumeration value="med"/>
 *     <xsd:enumeration value="lg"/>
 *   </xsd:restriction>
 * </xsd:simpleType>
 * ```
 *
 * @publicApi
 */
export const LineEndLength = {
    /** Small length */
    SMALL: "sm",
    /** Medium length */
    MEDIUM: "med",
    /** Large length */
    LARGE: "lg",
} as const;

/**
 * Options for line end (arrow) properties.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_LineEndProperties">
 *   <xsd:attribute name="type" type="ST_LineEndType" use="optional" default="none"/>
 *   <xsd:attribute name="w" type="ST_LineEndWidth" use="optional"/>
 *   <xsd:attribute name="len" type="ST_LineEndLength" use="optional"/>
 * </xsd:complexType>
 * ```
 */
export interface LineEndOptions {
    /** Arrow/head type */
    readonly type: keyof typeof LineEndType;
    /** Arrow width */
    readonly width?: keyof typeof LineEndWidth;
    /** Arrow length */
    readonly length?: keyof typeof LineEndLength;
}

/**
 * Creates a line end element (a:headEnd or a:tailEnd).
 *
 * @example
 * ```typescript
 * // Stealth arrow at start, medium size
 * createLineEnd("a:headEnd", { type: "STEALTH", width: "MEDIUM", length: "MEDIUM" });
 * // Triangle arrow at end
 * createLineEnd("a:tailEnd", { type: "TRIANGLE" });
 * ```
 */
export const createLineEnd = (name: string, options: LineEndOptions): XmlComponent =>
    new BuilderElement<{ readonly type?: string; readonly w?: string; readonly len?: string }>({
        attributes: {
            type: { key: "type", value: LineEndType[options.type] },
            w: { key: "w", value: options.width ? LineEndWidth[options.width] : undefined },
            len: { key: "len", value: options.length ? LineEndLength[options.length] : undefined },
        },
        name,
    });
