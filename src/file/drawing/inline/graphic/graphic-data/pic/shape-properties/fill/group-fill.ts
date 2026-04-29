/**
 * Group fill element for DrawingML shapes.
 *
 * This module provides group fill support which inherits the fill
 * from the parent group shape.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_GroupFillProperties
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

/**
 * Creates a group fill element (a:grpFill).
 *
 * This element specifies that the shape should inherit its fill
 * from the parent group shape. This is useful when shapes are
 * grouped together and should share the same fill.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_GroupFillProperties"/>
 * ```
 *
 * @example
 * ```typescript
 * // Shape inherits fill from parent group
 * createGroupFill();
 * ```
 */
export const createGroupFill = (): XmlComponent =>
    new BuilderElement({
        name: "a:grpFill",
    });
