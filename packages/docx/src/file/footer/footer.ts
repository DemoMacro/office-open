/**
 * Footer module for WordprocessingML documents.
 *
 * Footers are repeated at the bottom of each page in a section.
 *
 * Reference: http://officeopenxml.com/WPfooters.php
 *
 * @module
 */
import { InitializableXmlComponent } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import type { FileChild } from "../file-child";
import { FooterAttributes } from "./footer-attributes";

/**
 * Represents a footer in a WordprocessingML document.
 *
 * A footer is the portion of the document that appears at the bottom of each page in a section.
 * Footers can contain block-level elements such as paragraphs and tables. Each section can
 * have up to three different footers: first page, even pages, and odd pages.
 *
 * Reference: http://officeopenxml.com/WPfooters.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_HdrFtr">
 *   <xsd:group ref="EG_BlockLevelElts" minOccurs="1" maxOccurs="unbounded"/>
 * </xsd:complexType>
 *
 * <xsd:element name="ftr" type="CT_HdrFtr"/>
 * ```
 *
 * @example
 * ```typescript
 * // Create a simple footer with page numbers
 * const footer = new Footer(1);
 * footer.add(new Paragraph({
 *   alignment: AlignmentType.CENTER,
 *   children: [new TextRun("Page "), PageNumber.CURRENT]
 * }));
 *
 * // Create a footer with a table
 * const footer = new Footer(2);
 * footer.add(new Table({
 *   rows: [
 *     new TableRow({
 *       children: [new TableCell({ children: [new Paragraph("Footer Content")] })]
 *     })
 *   ]
 * }));
 * ```
 */
export class Footer extends InitializableXmlComponent {
    private readonly refId: number;

    public constructor(referenceNumber: number, initContent?: XmlComponent) {
        super("w:ftr", initContent);
        this.refId = referenceNumber;
        if (!initContent) {
            this.root.push(
                new FooterAttributes({
                    m: "http://schemas.openxmlformats.org/officeDocument/2006/math",
                    mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
                    o: "urn:schemas-microsoft-com:office:office",
                    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                    v: "urn:schemas-microsoft-com:vml",
                    w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                    w10: "urn:schemas-microsoft-com:office:word",
                    w14: "http://schemas.microsoft.com/office/word/2010/wordml",
                    w15: "http://schemas.microsoft.com/office/word/2012/wordml",
                    wne: "http://schemas.microsoft.com/office/word/2006/wordml",
                    wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                    wp14: "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
                    wpc: "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
                    wpg: "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
                    wpi: "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
                    wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
                }),
            );
        }
    }

    public get ReferenceId(): number {
        return this.refId;
    }

    public add(item: FileChild): void {
        this.root.push(item);
    }
}
