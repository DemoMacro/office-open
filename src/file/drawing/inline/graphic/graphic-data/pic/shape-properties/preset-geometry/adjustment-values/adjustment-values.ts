/**
 * Adjustment values module for preset geometries.
 *
 * This module provides adjustment value lists that can modify the appearance
 * of preset shape geometries.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_GeomGuideList, CT_GeomGuide
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

/**
 * A single geometry guide that defines an adjustment value.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_GeomGuide">
 *   <xsd:attribute name="name" type="ST_GeomGuideName" use="required"/>
 *   <xsd:attribute name="fmla" type="ST_GeomGuideFormula" use="required"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Adjust rounded rectangle corner radius
 * { name: "adj", formula: "val 16667" }
 *
 * // Adjust star inner radius
 * { name: "val", formula: "val 50000" }
 * ```
 *
 * Note: Formula examples like `val 16667` define fixed adjustment values.
 */
export interface GeometryGuide {
    /** Guide name (identifier) */
    readonly name: string;
    /** Guide formula (e.g., "val 16667") */
    readonly formula: string;
}

/**
 * Creates an adjustment values list element (a:avLst).
 *
 * The adjustment values list contains geometry guides that modify
 * the appearance of a preset geometric shape. When empty, default
 * values are used.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_GeomGuideList">
 *   <xsd:sequence>
 *     <xsd:element name="gd" type="CT_GeomGuide" minOccurs="0" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Empty adjustment values (defaults)
 * createAdjustmentValues();
 *
 * // With guides
 * createAdjustmentValues([
 *   { name: "adj", formula: "val 16667" },
 * ]);
 * ```
 */
export const createAdjustmentValues = (guides?: readonly GeometryGuide[]): XmlComponent => {
    const children: XmlComponent[] = [];

    if (guides) {
        for (const guide of guides) {
            children.push(
                new BuilderElement<{ readonly name: string; readonly fmla: string }>({
                    attributes: {
                        name: { key: "name", value: guide.name },
                        fmla: { key: "fmla", value: guide.formula },
                    },
                    name: "a:gd",
                }),
            );
        }
    }

    return new BuilderElement({
        children,
        name: "a:avLst",
    });
};
