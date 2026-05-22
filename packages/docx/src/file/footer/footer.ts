/**
 * Footer module for WordprocessingML documents.
 *
 * Footers are repeated at the bottom of each page in a section.
 *
 * Reference: http://officeopenxml.com/WPfooters.php
 *
 * @module
 */
import type { XmlComponent } from "@file/xml-components";

import { HeaderFooterBase, FOOTER_NAMESPACES } from "../header-footer-base";

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
export class Footer extends HeaderFooterBase {
  public constructor(referenceNumber: number, initContent?: XmlComponent) {
    super("w:ftr", referenceNumber, FOOTER_NAMESPACES, initContent);
  }
}
