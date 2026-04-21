/**
 * Document Properties module for DrawingML elements.
 *
 * This module provides non-visual properties for drawing elements,
 * including name, description, accessibility information, and hyperlinks.
 *
 * Reference: https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_docPr_topic_ID0ES32OB.html
 *
 * @module
 */
import { ConcreteHyperlink } from "@file/paragraph";
import { TargetModeType } from "@file/relationships/relationship/relationship";
import { NextAttributeComponent, XmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";
import { docPropertiesUniqueNumericIdGen, uniqueId } from "@util/convenience-functions";

import { createHyperlinkClick, createHyperlinkHover } from "./doc-properties-children";

// <complexType name="CT_NonVisualDrawingProps">
//     <sequence>
//         <element name="hlinkClick" type="CT_Hyperlink" minOccurs="0" maxOccurs="1" />
//         <element name="hlinkHover" type="CT_Hyperlink" minOccurs="0" maxOccurs="1" />
//         <element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1" />
//     </sequence>
//     <attribute name="id" type="ST_DrawingElementId" use="required" />
//     <attribute name="name" type="xsd:string" use="required" />
//     <attribute name="descr" type="xsd:string" use="optional" default="" />
//     <attribute name="hidden" type="xsd:boolean" use="optional" default="false" />
// </complexType>

/**
 * Options for hyperlinks on a drawing element.
 */
export interface HyperlinkOptions {
    /** URL for click hyperlink */
    readonly click?: string;
    /** URL for hover hyperlink */
    readonly hover?: string;
}

/**
 * Options for configuring document properties of a drawing.
 *
 * @see {@link DocProperties}
 */
export interface DocPropertiesOptions {
    /** Name of the drawing element (used for identification) */
    readonly name: string;
    /** Description/alt text for accessibility */
    readonly description?: string;
    /** Title of the drawing element */
    readonly title?: string;
    readonly id?: string;
    /** Hyperlink options for click and hover actions */
    readonly hyperlink?: HyperlinkOptions;
}

const HYPERLINK_RELATIONSHIP_TYPE =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

/**
 * Represents non-visual drawing properties in a WordprocessingML document.
 *
 * DocProperties contains metadata about a drawing element such as
 * its name, description (alt text), and title for accessibility.
 *
 * Reference: https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_docPr_topic_ID0ES32OB.html
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_NonVisualDrawingProps">
 *   <xsd:sequence>
 *     <xsd:element name="hlinkClick" type="CT_Hyperlink" minOccurs="0"/>
 *     <xsd:element name="hlinkHover" type="CT_Hyperlink" minOccurs="0"/>
 *     <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="id" type="ST_DrawingElementId" use="required"/>
 *   <xsd:attribute name="name" type="xsd:string" use="required"/>
 *   <xsd:attribute name="descr" type="xsd:string"/>
 *   <xsd:attribute name="hidden" type="xsd:boolean"/>
 * </xsd:complexType>
 * ```
 */
export class DocProperties extends XmlComponent {
    private readonly docPropertiesUniqueNumericId = docPropertiesUniqueNumericIdGen();
    private readonly hyperlink?: HyperlinkOptions;

    public constructor(
        { name, description, title, id, hyperlink }: DocPropertiesOptions = {
            description: "",
            name: "",
            title: "",
        },
    ) {
        super("wp:docPr");

        this.hyperlink = hyperlink;

        const attributes: Record<
            string,
            { readonly key: string; readonly value: string | number }
        > = {
            id: {
                key: "id",
                value: id ?? this.docPropertiesUniqueNumericId(),
            },
            name: {
                key: "name",
                value: name,
            },
        };

        if (description !== null && description !== undefined) {
            attributes.description = {
                key: "descr",
                value: description,
            };
        }

        if (title !== null && title !== undefined) {
            attributes.title = {
                key: "title",
                value: title,
            };
        }

        this.root.push(new NextAttributeComponent(attributes));
    }

    public prepForXml(context: IContext): IXmlableObject | undefined {
        // Stack-based detection (backward compatible with ConcreteHyperlink wrapping)
        let hasStackClick = false;
        for (let i = context.stack.length - 1; i >= 0; i--) {
            const element = context.stack[i];
            if (!(element instanceof ConcreteHyperlink)) {
                continue;
            }

            this.root.push(createHyperlinkClick(element.linkId, true));
            hasStackClick = true;
            break;
        }

        // Explicit hyperlink options take precedence over stack-based detection
        if (this.hyperlink) {
            if (this.hyperlink.click && !hasStackClick) {
                const linkId = uniqueId();
                context.viewWrapper.Relationships.addRelationship(
                    linkId,
                    HYPERLINK_RELATIONSHIP_TYPE,
                    this.hyperlink.click,
                    TargetModeType.EXTERNAL,
                );
                this.root.push(createHyperlinkClick(linkId, true));
            }

            if (this.hyperlink.hover) {
                const linkId = uniqueId();
                context.viewWrapper.Relationships.addRelationship(
                    linkId,
                    HYPERLINK_RELATIONSHIP_TYPE,
                    this.hyperlink.hover,
                    TargetModeType.EXTERNAL,
                );
                this.root.push(createHyperlinkHover(linkId, true));
            }
        }

        return super.prepForXml(context);
    }
}
