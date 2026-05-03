/**
 * Preset geometry module for DrawingML shapes.
 *
 * This module provides predefined shape geometries that can be applied
 * to pictures and shapes without requiring custom path definitions.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_PresetGeometry2D
 *
 * @module
 */
import { XmlComponent } from "../../xml-components";
import { createAdjustmentValues } from "./adjustment-values";
import type { GeometryGuide } from "./adjustment-values";
import { PresetGeometryAttributes } from "./preset-geometry-attributes";

/**
 * Options for preset geometry.
 */
export interface PresetGeometryOptions {
    /**
     * Preset shape type (ST_ShapeType).
     *
     * Defaults to `"rect"` if not specified.
     */
    readonly preset?: string;
    /**
     * Adjustment values that modify the base shape.
     *
     * Each guide has a name and formula (e.g., `{ name: "adj", formula: "val 16667" }`).
     */
    readonly adjustmentValues?: readonly GeometryGuide[];
}

/**
 * Represents a preset geometry for a DrawingML shape.
 *
 * This element specifies when a preset geometric shape should be used instead
 * of a custom geometry. It includes a shape preset identifier and optional
 * adjustment values that modify the base shape.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_PresetGeometry2D">
 *   <xsd:sequence>
 *     <xsd:element name="avLst" type="CT_GeomGuideList" minOccurs="0" maxOccurs="1"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="prst" type="ST_ShapeType" use="required"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Default rectangle
 * const geometry = new PresetGeometry();
 *
 * // Rounded rectangle with adjustment
 * const geometry = new PresetGeometry({
 *   preset: "roundRect",
 *   adjustmentValues: [{ name: "adj", formula: "val 16667" }],
 * });
 * ```
 */
export class PresetGeometry extends XmlComponent {
    public constructor(options?: PresetGeometryOptions) {
        super("a:prstGeom");

        this.root.push(
            new PresetGeometryAttributes({
                prst: options?.preset ?? "rect",
            }),
        );

        this.root.push(createAdjustmentValues(options?.adjustmentValues));
    }
}
