/**
 * Custom dash pattern for DrawingML outlines.
 *
 * This module provides support for custom dash patterns defined by
 * a list of dash stops, each specifying dash and space lengths as
 * percentages relative to line width.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_DashStopList, CT_DashStop
 *
 * @module
 */
import { BuilderElement } from "../../xml-components";
import type { XmlComponent } from "../../xml-components";

/**
 * A single dash stop in a custom dash pattern.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_DashStop">
 *   <xsd:attribute name="d" type="ST_PositivePercentage" use="required"/>
 *   <xsd:attribute name="sp" type="ST_PositivePercentage" use="required"/>
 * </xsd:complexType>
 * ```
 *
 * `ST_PositivePercentage` is a string matching `[0-9]+(\.[0-9]+)?%`
 * (e.g., `"100%"`, `"500%"`, `"50.5%"`).
 *
 * @example
 * ```typescript
 * // A dash stop with 500% dash length and 200% space length
 * { d: "500%", sp: "200%" }
 * ```
 */
export interface DashStop {
    /** Dash length as a percentage of line width (e.g., "500%") */
    readonly d: string;
    /** Space length as a percentage of line width (e.g., "200%") */
    readonly sp: string;
}

/**
 * Creates a custom dash element (a:custDash) for outlines.
 *
 * The custom dash element specifies a repeating pattern of dash and space
 * segments. Each segment is defined by a `DashStop` with dash (`d`) and
 * space (`sp`) lengths as positive percentages of the line width.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_DashStopList">
 *   <xsd:sequence>
 *     <xsd:element name="ds" type="CT_DashStop" minOccurs="0" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Simple dash pattern
 * createCustomDash([
 *   { d: "500%", sp: "200%" },
 * ]);
 *
 * // Complex alternating pattern
 * createCustomDash([
 *   { d: "500%", sp: "200%" },
 *   { d: "100%", sp: "200%" },
 * ]);
 * ```
 */
export const createCustomDash = (stops: readonly DashStop[]): XmlComponent => {
    const children: XmlComponent[] = [];

    for (const stop of stops) {
        children.push(
            new BuilderElement<{ readonly d: string; readonly sp: string }>({
                attributes: {
                    d: { key: "d", value: stop.d },
                    sp: { key: "sp", value: stop.sp },
                },
                name: "a:ds",
            }),
        );
    }

    return new BuilderElement({
        children,
        name: "a:custDash",
    });
};
