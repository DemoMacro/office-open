/**
 * Footnote module for WordprocessingML documents.
 *
 * This module provides support for footnotes that appear at the bottom
 * of the page referenced from the main document text.
 *
 * Reference: http://officeopenxml.com/WPfootnotes.php
 *
 * @module
 */
import type { FileChild } from "@file/file-child";
import { XmlComponent } from "@file/xml-components";

import { FootnoteRefRun } from "./run/footnote-ref-run";

/**
 * Enumeration of footnote types.
 *
 * Reference: http://officeopenxml.com/WPfootnotes.php
 *
 * @publicApi
 */
export const FootnoteType = {
  /** Separator line between body text and footnotes */
  SEPERATOR: "separator",
  /** Continuation separator for footnotes spanning pages */
  CONTINUATION_SEPERATOR: "continuationSeparator",
} as const;

/**
 * Options for creating a Footnote.
 *
 * @property id - Unique numeric identifier for this footnote
 * @property type - Type of footnote (separator, continuationSeparator, or normal)
 * @property children - Array of paragraphs that make up the footnote content
 *
 * @see {@link Footnote}
 */
export interface FootnoteOptions {
  /** Unique numeric identifier for this footnote */
  readonly id: number;
  /** Type of footnote (separator or continuation separator) */
  readonly type?: (typeof FootnoteType)[keyof typeof FootnoteType];
  /** Content of the footnote (paragraphs, tables, etc.) */
  readonly children: readonly FileChild[];
}

/**
 * Represents a footnote in a WordprocessingML document.
 *
 * Footnotes appear at the bottom of the page and are referenced
 * from the main document text using footnote reference marks.
 *
 * Reference: http://officeopenxml.com/WPfootnotes.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_FtnEdn">
 *   <xsd:sequence>
 *     <xsd:group ref="EG_BlockLevelElts" minOccurs="0" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="type" type="ST_FtnEdn"/>
 *   <xsd:attribute name="id" type="ST_DecimalNumber" use="required"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Create a footnote with content
 * const footnote = new Footnote({
 *   id: 1,
 *   children: [
 *     new Paragraph({
 *       children: [
 *         new TextRun("This is the footnote content."),
 *       ],
 *     }),
 *   ],
 * });
 * ```
 */
export class Footnote extends XmlComponent {
  public constructor(options: FootnoteOptions) {
    super("w:footnote");
    const attr: Record<string, string | number> = { "w:id": options.id };
    if (options.type !== undefined) {
      attr["w:type"] = options.type;
    }
    this.root.push({ _attr: attr });

    for (let i = 0; i < options.children.length; i++) {
      const child = options.children[i];

      if (i === 0 && "addRunToFront" in child) {
        (child as { addRunToFront: (run: unknown) => void }).addRunToFront(new FootnoteRefRun());
      }

      this.root.push(child);
    }
  }
}
