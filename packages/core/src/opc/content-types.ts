/**
 * Content Types components for OPC (Open Packaging Convention).
 *
 * Provides Default and Override elements used in [Content_Types].xml
 * to map file extensions and part paths to MIME content types.
 *
 * Reference: ECMA-376 Part 2, Section 10.1
 *
 * @module
 */
import { BuilderElement, XmlAttributeComponent } from "../xml-components";
import type { XmlComponent } from "../xml-components";

export interface DefaultAttributes {
    readonly contentType: string;
    readonly extension?: string;
}

/**
 * Creates a Default element mapping a file extension to a MIME content type.
 */
export const createDefault = (contentType: string, extension?: string): XmlComponent =>
    new BuilderElement<DefaultAttributes>({
        attributes: {
            contentType: { key: "ContentType", value: contentType },
            extension: { key: "Extension", value: extension },
        },
        name: "Default",
    });

export interface OverrideAttributes {
    readonly contentType: string;
    readonly partName?: string;
}

/**
 * Creates an Override element mapping a specific part path to a MIME content type.
 */
export const createOverride = (contentType: string, partName?: string): XmlComponent =>
    new BuilderElement<OverrideAttributes>({
        attributes: {
            contentType: { key: "ContentType", value: contentType },
            partName: { key: "PartName", value: partName },
        },
        name: "Override",
    });

/**
 * Attributes for the Types (Content Types) root element.
 * Declares the XML namespace for the content types part.
 */
export class ContentTypeAttributes extends XmlAttributeComponent<{
    readonly xmlns?: string;
}> {
    protected readonly xmlKeys = { xmlns: "xmlns" };
}
