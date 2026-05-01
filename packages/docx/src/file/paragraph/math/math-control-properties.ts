/**
 * Math Control Properties module for Office MathML.
 *
 * This module provides the control properties element (m:ctrlPr)
 * used in various math structures.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_ctrlPr-1.html
 *
 * @module
 */
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

/**
 * Options for math control properties.
 *
 * @see {@link createMathControlProperties}
 */
export interface MathControlPropertiesOptions {
    /** Insertion tracking reference */
    readonly insertionReference?: string;
    /** Deletion tracking reference */
    readonly deletionReference?: string;
}

/**
 * Creates math control properties element (m:ctrlPr).
 *
 * This element specifies control properties for math elements,
 * including tracking changes for insertions and deletions.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-m_ctrlPr-1.html
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_CtrlPr">
 *   <xsd:sequence>
 *     <xsd:group ref="w:EG_RPrMath" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 */
export const createMathControlProperties = (
    options?: MathControlPropertiesOptions,
): XmlComponent => {
    const children: XmlComponent[] = [];

    if (options?.insertionReference !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { id: { key: "w:id", value: options.insertionReference } },
                name: "w:ins",
            }),
        );
    }

    if (options?.deletionReference !== undefined) {
        children.push(
            new BuilderElement({
                attributes: { id: { key: "w:id", value: options.deletionReference } },
                name: "w:del",
            }),
        );
    }

    return new BuilderElement({
        children,
        name: "m:ctrlPr",
    });
};
