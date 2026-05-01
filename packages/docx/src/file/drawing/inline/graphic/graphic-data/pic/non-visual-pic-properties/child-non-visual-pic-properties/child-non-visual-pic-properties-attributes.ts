/**
 * Attributes for non-visual picture properties.
 *
 * @module
 */
import { XmlAttributeComponent } from "@file/xml-components";

/**
 * Attributes for CT_NonVisualPictureProperties.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:attribute name="preferRelativeResize" type="xsd:boolean" use="optional" default="true"/>
 * ```
 *
 * @internal
 */
export class ChildNonVisualPropertiesAttributes extends XmlAttributeComponent<{
    readonly preferRelativeResize?: boolean;
}> {
    protected readonly xmlKeys = {
        preferRelativeResize: "preferRelativeResize",
    };
}
