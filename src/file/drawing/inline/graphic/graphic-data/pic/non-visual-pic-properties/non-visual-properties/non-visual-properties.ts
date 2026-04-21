/**
 * Non-visual drawing properties module.
 *
 * This module provides basic metadata for drawing elements including
 * ID, name, description, and hyperlink support.
 *
 * Reference: http://officeopenxml.com/drwPic.php
 *
 * @module
 */
import type { HyperlinkOptions } from "@file/drawing/doc-properties/doc-properties";
import {
    createHyperlinkClick,
    createHyperlinkHover,
} from "@file/drawing/doc-properties/doc-properties-children";
import { ConcreteHyperlink } from "@file/paragraph";
import { TargetModeType } from "@file/relationships/relationship/relationship";
import { XmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";
import { uniqueId } from "@util/convenience-functions";

import { NonVisualPropertiesAttributes } from "./non-visual-properties-attributes";

const HYPERLINK_RELATIONSHIP_TYPE =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

/**
 * Represents non-visual drawing properties for pictures.
 *
 * This element specifies non-visual properties for a DrawingML object.
 * These include identification, naming, description, and hyperlink properties.
 *
 * Reference: http://officeopenxml.com/drwPic.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_NonVisualDrawingProps">
 *   <xsd:sequence>
 *     <xsd:element name="hlinkClick" type="CT_Hyperlink" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="hlinkHover" type="CT_Hyperlink" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="id" type="ST_DrawingElementId" use="required"/>
 *   <xsd:attribute name="name" type="xsd:string" use="required"/>
 *   <xsd:attribute name="descr" type="xsd:string" use="optional" default=""/>
 *   <xsd:attribute name="hidden" type="xsd:boolean" use="optional" default="false"/>
 *   <xsd:attribute name="title" type="xsd:string" use="optional" default=""/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * const cNvPr = new NonVisualProperties();
 * ```
 */
export class NonVisualProperties extends XmlComponent {
    private readonly hyperlink?: HyperlinkOptions;

    public constructor(hyperlink?: HyperlinkOptions) {
        super("pic:cNvPr");

        this.hyperlink = hyperlink;

        this.root.push(
            new NonVisualPropertiesAttributes({
                descr: "",
                id: 0,
                name: "",
            }),
        );
    }

    public prepForXml(context: IContext): IXmlableObject | undefined {
        // Stack-based detection (backward compatible with ConcreteHyperlink wrapping)
        let hasStackClick = false;
        for (let i = context.stack.length - 1; i >= 0; i--) {
            const element = context.stack[i];
            if (!(element instanceof ConcreteHyperlink)) {
                continue;
            }

            this.root.push(createHyperlinkClick(element.linkId, false));
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
                this.root.push(createHyperlinkClick(linkId, false));
            }

            if (this.hyperlink.hover) {
                const linkId = uniqueId();
                context.viewWrapper.Relationships.addRelationship(
                    linkId,
                    HYPERLINK_RELATIONSHIP_TYPE,
                    this.hyperlink.hover,
                    TargetModeType.EXTERNAL,
                );
                this.root.push(createHyperlinkHover(linkId, false));
            }
        }

        return super.prepForXml(context);
    }
}
