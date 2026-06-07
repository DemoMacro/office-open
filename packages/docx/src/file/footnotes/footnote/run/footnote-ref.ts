/**
 * Footnote reference marker module for WordprocessingML documents.
 *
 * This module provides the automatic footnote reference marker that appears
 * at the beginning of footnote content.
 *
 * Reference: http://officeopenxml.com/WPfootnotes.php
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

/**
 * Represents a footnote reference marker in WordprocessingML.
 *
 * FootnoteRef creates the automatic reference mark (typically a superscript
 * number) that appears at the beginning of the footnote content itself.
 * This corresponds to the w:footnoteRef element in OOXML.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_FtnEdnRef">
 *   <xsd:attribute name="customMarkFollows" type="s:ST_OnOff" use="optional"/>
 *   <xsd:attribute name="id" use="required" type="ST_DecimalNumber"/>
 * </xsd:complexType>
 * ```
 *
 * @internal
 */
export class FootnoteRef extends XmlComponent {
  public constructor(options?: { customMarkFollows?: boolean }) {
    super("w:footnoteRef");
    if (options?.customMarkFollows !== undefined) {
      this.root.push({ _attr: { "w:customMarkFollows": options.customMarkFollows ? "1" : "0" } });
    }
  }
}
