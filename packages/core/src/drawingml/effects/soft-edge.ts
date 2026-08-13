/**
 * Soft edge effect for DrawingML shapes.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_SoftEdgesEffect
 *
 * @module
 */
import { convertToEmu } from "../../util/converters";
import type { UniversalMeasure } from "../../util/values";

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
 * @param rad - Soft edge radius in EMUs (number) or UniversalMeasure (required)
 */
export const createSoftEdgeEffect = (rad: number | UniversalMeasure): string =>
  `<a:softEdge rad="${convertToEmu(rad)}"/>`;
