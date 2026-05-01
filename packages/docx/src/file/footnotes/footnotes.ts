/**
 * Footnotes module for WordprocessingML documents.
 *
 * This module manages the collection of footnotes in a document, including
 * separator and continuation separator footnotes that appear by default.
 *
 * Reference: http://officeopenxml.com/WPfootnotes.php
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

import { LineRuleType, Paragraph } from "../paragraph";
import { Footnote, FootnoteType } from "./footnote/footnote";
import { ContinuationSeperatorRun } from "./footnote/run/continuation-seperator-run";
import { SeperatorRun } from "./footnote/run/seperator-run";
import { FootnotesAttributes } from "./footnotes-attributes";

/**
 * Represents the footnotes collection in a WordprocessingML document.
 *
 * FootNotes manages all footnotes in a document and automatically creates
 * the required separator and continuation separator footnotes. These special
 * footnotes define the line that separates body text from footnotes and the
 * line used when footnotes continue across pages.
 *
 * Reference: http://officeopenxml.com/WPfootnotes.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Footnotes">
 *   <xsd:sequence maxOccurs="unbounded">
 *     <xsd:element name="footnote" type="CT_FtnEdn" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // FootNotes is typically managed internally by the Document class
 * // Users create footnotes through the Document API
 * const doc = new Document({
 *   sections: [{
 *     children: [
 *       new Paragraph({
 *         children: [
 *           new TextRun("Text with footnote"),
 *           new FootnoteReferenceRun(1),
 *         ],
 *       }),
 *     ],
 *     footnotes: {
 *       1: {
 *         children: [new Paragraph("Footnote content")],
 *       },
 *     },
 *   }],
 * });
 * ```
 */
export class FootNotes extends XmlComponent {
    public constructor() {
        super("w:footnotes");

        this.root.push(
            new FootnotesAttributes({
                Ignorable: "w14 w15 wp14",
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

        const begin = new Footnote({
            children: [
                new Paragraph({
                    children: [new SeperatorRun()],
                    spacing: {
                        after: 0,
                        line: 240,
                        lineRule: LineRuleType.AUTO,
                    },
                }),
            ],
            id: -1,
            type: FootnoteType.SEPERATOR,
        });

        this.root.push(begin);

        const spacing = new Footnote({
            children: [
                new Paragraph({
                    children: [new ContinuationSeperatorRun()],
                    spacing: {
                        after: 0,
                        line: 240,
                        lineRule: LineRuleType.AUTO,
                    },
                }),
            ],
            id: 0,
            type: FootnoteType.CONTINUATION_SEPERATOR,
        });

        this.root.push(spacing);
    }

    /**
     * Creates and adds a new footnote to the collection.
     *
     * @param id - Unique numeric identifier for the footnote
     * @param paragraph - Array of paragraphs that make up the footnote content
     */
    public createFootNote(id: number, paragraph: readonly Paragraph[]): void {
        const footnote = new Footnote({
            children: paragraph,
            id: id,
        });

        this.root.push(footnote);
    }
}
