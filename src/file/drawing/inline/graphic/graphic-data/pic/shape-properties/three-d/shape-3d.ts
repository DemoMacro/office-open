/**
 * 3D shape properties for DrawingML.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_Shape3D
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import { createColorElement } from "../outline/solid-fill";
import type { SolidFillOptions } from "../outline/solid-fill";
import { createBevel, createBottomBevel } from "./bevel";
import type { BevelOptions } from "./bevel";

/**
 * Preset material types for 3D shapes (15 variations).
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, ST_PresetMaterialType
 */
export const PresetMaterialType = {
    LEGACY_MATTE: "legacyMatte",
    LEGACY_PLASTIC: "legacyPlastic",
    LEGACY_METAL: "legacyMetal",
    LEGACY_WIREFRAME: "legacyWireframe",
    MATTE: "matte",
    PLASTIC: "plastic",
    METAL: "metal",
    WARM_MATTE: "warmMatte",
    TRANSLUCENT_POWDER: "translucentPowder",
    POWDER: "powder",
    DK_EDGE: "dkEdge",
    SOFT_EDGE: "softEdge",
    CLEAR: "clear",
    FLAT: "flat",
    SOFT_METAL: "softmetal",
} as const;

/**
 * Options for 3D shape properties.
 */
export interface Shape3DOptions {
    /** Top bevel options */
    readonly bevelT?: BevelOptions;
    /** Bottom bevel options */
    readonly bevelB?: BevelOptions;
    /** Extrusion color (CT_Color / EG_ColorChoice) */
    readonly extrusionColor?: SolidFillOptions;
    /** Contour color (CT_Color / EG_ColorChoice) */
    readonly contourColor?: SolidFillOptions;
    /** Depth in EMUs (default 0) */
    readonly z?: number;
    /** Extrusion height in EMUs (default 0) */
    readonly extrusionH?: number;
    /** Contour width in EMUs (default 0) */
    readonly contourW?: number;
    /** Material preset type */
    readonly prstMaterial?: keyof typeof PresetMaterialType;
}

/**
 * Creates a 3D shape properties element (a:sp3d).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Shape3D">
 *   <xsd:sequence>
 *     <xsd:element name="bevelT" type="CT_Bevel" minOccurs="0"/>
 *     <xsd:element name="bevelB" type="CT_Bevel" minOccurs="0"/>
 *     <xsd:element name="extrusionClr" type="CT_Color" minOccurs="0"/>
 *     <xsd:element name="contourClr" type="CT_Color" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="z" type="ST_Coordinate" default="0"/>
 *   <xsd:attribute name="extrusionH" type="ST_PositiveCoordinate" default="0"/>
 *   <xsd:attribute name="contourW" type="ST_PositiveCoordinate" default="0"/>
 *   <xsd:attribute name="prstMaterial" type="ST_PresetMaterialType" default="warmMatte"/>
 * </xsd:complexType>
 * ```
 */
export const createShape3D = (options: Shape3DOptions): XmlComponent => {
    const children: XmlComponent[] = [];

    if (options.bevelT) {
        children.push(createBevel(options.bevelT));
    }
    if (options.bevelB) {
        children.push(createBottomBevel(options.bevelB));
    }
    if (options.extrusionColor) {
        children.push(
            new BuilderElement({
                children: [createColorElement(options.extrusionColor)],
                name: "a:extrusionClr",
            }),
        );
    }
    if (options.contourColor) {
        children.push(
            new BuilderElement({
                children: [createColorElement(options.contourColor)],
                name: "a:contourClr",
            }),
        );
    }

    const hasAttributes =
        options.z !== undefined ||
        options.extrusionH !== undefined ||
        options.contourW !== undefined ||
        options.prstMaterial !== undefined;

    const attributePayload = hasAttributes
        ? {
              ...(options.z !== undefined && { z: { key: "z", value: options.z } }),
              ...(options.extrusionH !== undefined && {
                  extrusionH: { key: "extrusionH", value: options.extrusionH },
              }),
              ...(options.contourW !== undefined && {
                  contourW: { key: "contourW", value: options.contourW },
              }),
              ...(options.prstMaterial !== undefined && {
                  prstMaterial: {
                      key: "prstMaterial",
                      value: PresetMaterialType[options.prstMaterial],
                  },
              }),
          }
        : undefined;

    return new BuilderElement({
        attributes: attributePayload as never,
        children,
        name: "a:sp3d",
    });
};
