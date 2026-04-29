/**
 * Structured Document Tag (SDT) module for WordprocessingML documents.
 *
 * SDTs (content controls) are containers that wrap content with metadata and
 * behavior. They exist at multiple levels: inline (run), block, row, and cell.
 *
 * This module provides:
 * - `StructuredDocumentTagRun` — inline SDT (CT_SdtRun) for use within paragraphs
 * - Re-exports of SDT property types from the table-of-contents module
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_SdtRun, CT_SdtBlock
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

import {
    StructuredDocumentTagContent,
    StructuredDocumentTagProperties,
} from "../table-of-contents";
import type { SdtPropertiesOptions } from "../table-of-contents";

/**
 * Options for creating an inline Structured Document Tag (CT_SdtRun).
 */
export interface ISdtRunOptions {
    /** SDT properties (alias, tag, type discriminator, etc.) */
    readonly properties: SdtPropertiesOptions;
    /** Content children (runs, text runs, etc.) to place inside the SDT */
    readonly children?: readonly XmlComponent[];
}

/**
 * An inline Structured Document Tag (CT_SdtRun).
 *
 * Represents a content control that appears within a paragraph, alongside
 * runs and other inline elements. The SDT wraps its content in `w:sdtContent`.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_SdtRun">
 *   <xsd:sequence>
 *     <xsd:element name="sdtPr" type="CT_SdtPr" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="sdtEndPr" type="CT_SdtEndPr" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="sdtContent" type="CT_SdtContentRun" minOccurs="0" maxOccurs="1"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Plain text content control
 * new StructuredDocumentTagRun({
 *   properties: { text: { multiLine: true }, alias: "Description" },
 *   children: [new TextRun("Default text")],
 * });
 *
 * // ComboBox content control
 * new StructuredDocumentTagRun({
 *   properties: {
 *     alias: "Color",
 *     comboBox: { items: [{ displayText: "Red" }, { displayText: "Blue" }] },
 *   },
 *   children: [new TextRun("Red")],
 * });
 * ```
 */
export class StructuredDocumentTagRun extends XmlComponent {
    public constructor(options: ISdtRunOptions) {
        super("w:sdt");
        this.root.push(new StructuredDocumentTagProperties(options.properties));
        if (options.children && options.children.length > 0) {
            const content = new StructuredDocumentTagContent();
            for (const child of options.children) {
                content.addChildElement(child);
            }
            this.root.push(content);
        }
    }
}

/**
 * Options for creating a block-level Structured Document Tag (CT_SdtBlock).
 */
export interface ISdtBlockOptions {
    /** SDT properties */
    readonly properties: SdtPropertiesOptions;
    /** Content children (paragraphs, tables, etc.) to place inside the SDT */
    readonly children?: readonly XmlComponent[];
}

/**
 * A block-level Structured Document Tag (CT_SdtBlock).
 *
 * Represents a content control at the block level (alongside paragraphs and tables).
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_SdtBlock">
 *   <xsd:sequence>
 *     <xsd:element name="sdtPr" type="CT_SdtPr" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="sdtEndPr" type="CT_SdtEndPr" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="sdtContent" type="CT_SdtContentBlock" minOccurs="0" maxOccurs="1"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Block-level rich text SDT
 * new StructuredDocumentTagBlock({
 *   properties: { richText: true, alias: "Block Content" },
 *   children: [new Paragraph({ text: "Content here" })],
 * });
 * ```
 */
export class StructuredDocumentTagBlock extends XmlComponent {
    public constructor(options: ISdtBlockOptions) {
        super("w:sdt");
        this.root.push(new StructuredDocumentTagProperties(options.properties));
        if (options.children && options.children.length > 0) {
            const content = new StructuredDocumentTagContent();
            for (const child of options.children) {
                content.addChildElement(child);
            }
            this.root.push(content);
        }
    }
}
