/**
 * Soft edge effect for DrawingML shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_SoftEdgesEffect
 *
 * @module
 */
import { BuilderElement } from "../../xml-components";

/**
 * Creates a soft edge effect element.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_SoftEdgesEffect">
 *   <xsd:attribute name="rad" type="ST_PositiveCoordinate" use="required"/>
 * </xsd:complexType>
 * ```
 *
 * @param rad - Soft edge radius in EMUs (required)
 */
export const createSoftEdgeEffect = (rad: number) =>
    new BuilderElement<{ readonly rad: number }>({
        attributes: {
            rad: { key: "rad", value: rad },
        },
        name: "a:softEdge",
    });
