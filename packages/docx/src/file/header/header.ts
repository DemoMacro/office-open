/**
 * Header module for WordprocessingML documents.
 *
 * Headers are repeated at the top of each page in a section.
 *
 * Reference: http://officeopenxml.com/WPheaders.php
 *
 * @module
 */
import type { XmlComponent } from "@file/xml-components";

import { HeaderFooterBase, HEADER_NAMESPACES } from "../header-footer-base";

/**
 * Represents a header in a WordprocessingML document.
 *
 * A header is the portion of the document that appears at the top of each page in a section.
 * Headers can contain block-level elements such as paragraphs and tables. Each section can
 * have up to three different headers: first page, even pages, and odd pages.
 *
 * Reference: http://officeopenxml.com/WPheaders.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_HdrFtr">
 *   <xsd:group ref="EG_BlockLevelElts" minOccurs="1" maxOccurs="unbounded"/>
 * </xsd:complexType>
 *
 * <xsd:element name="hdr" type="CT_HdrFtr"/>
 * ```
 *
 * @example
 * ```typescript
 * // Create a simple header
 * const header = new Header(1);
 * header.add(new Paragraph("Company Name"));
 *
 * // Create a header with a table
 * const header = new Header(2);
 * header.add(new Table({
 *   rows: [
 *     new TableRow({
 *       children: [new TableCell({ children: [new Paragraph("Header Content")] })]
 *     })
 *   ]
 * }));
 * ```
 */
export class Header extends HeaderFooterBase {
  public constructor(referenceNumber: number, initContent?: XmlComponent) {
    super("w:hdr", referenceNumber, HEADER_NAMESPACES, initContent);
  }
}
