/**
 * Numbering properties module for WordprocessingML documents.
 *
 * This module provides numbering and list properties for paragraphs.
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";
import type { IXmlableObject } from "@file/xml-components";

/**
 * Represents numbering properties for a paragraph.
 *
 * The numPr element specifies the numbering definition instance and level
 * for the paragraph, enabling numbered and bulleted lists.
 *
 * Reference: http://officeopenxml.com/WPnumbering.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_NumPr">
 *   <xsd:sequence>
 *     <xsd:element name="ilvl" type="CT_DecimalNumber" minOccurs="0"/>
 *     <xsd:element name="numId" type="CT_DecimalNumber" minOccurs="0"/>
 *     <xsd:element name="numberingChange" type="CT_TrackChangeNumbering" minOccurs="0"/>
 *     <xsd:element name="ins" type="CT_TrackChange" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Create a bulleted list item at level 0
 * new Paragraph({
 *   numbering: {
 *     reference: "my-bullet-list",
 *     level: 0,
 *   },
 *   children: [new TextRun("First item")],
 * });
 *
 * // Create a numbered list item at level 1
 * new Paragraph({
 *   numbering: {
 *     reference: "my-numbered-list",
 *     level: 1,
 *   },
 *   children: [new TextRun("Nested item")],
 * });
 * ```
 */
/**
 * Build numbering properties (w:numPr) as IXmlableObject without allocating XmlComponent tree.
 */
export function buildNumberProperties(
  numberId: number | string,
  indentLevel: number,
  numberingChange?: {
    original: string;
    id: string;
    author: string;
    date?: string;
  },
): IXmlableObject {
  const children: IXmlableObject[] = [
    { "w:ilvl": { _attr: { "w:val": Math.min(indentLevel, 9) } } },
    {
      "w:numId": {
        _attr: { "w:val": typeof numberId === "string" ? `{${numberId}}` : numberId },
      },
    },
  ];

  if (numberingChange) {
    const changeAttrs: Record<string, string> = {
      "w:original": numberingChange.original,
      "w:id": numberingChange.id,
      "w:author": numberingChange.author,
    };
    if (numberingChange.date !== undefined) changeAttrs["w:date"] = numberingChange.date;
    children.push({ "w:numberingChange": { _attr: changeAttrs } });
  }

  return { "w:numPr": children };
}

export class NumberProperties extends XmlComponent {
  public constructor(
    numberId: number | string,
    indentLevel: number,
    numberingChange?: {
      original: string;
      id: string;
      author: string;
      date?: string;
    },
  ) {
    super("w:numPr");
    this.root.push(new IndentLevel(indentLevel));
    this.root.push(new NumberId(numberId));
    if (numberingChange) {
      this.root.push(new NumberingChange(numberingChange));
    }
  }
}

/**
 * Represents the indentation level (ilvl) for a numbered or bulleted list.
 *
 * The ilvl element specifies the list level (0-9) for the paragraph.
 *
 * @internal
 */
class IndentLevel extends XmlComponent {
  public constructor(level: number) {
    super("w:ilvl");

    this.root.push({ _attr: { "w:val": Math.min(level, 9) } });
  }
}

/**
 * Represents the numbering definition ID (numId) for a numbered or bulleted list.
 *
 * The numId element specifies which numbering definition to use for the paragraph.
 *
 * @internal
 */
class NumberId extends XmlComponent {
  public constructor(id: number | string) {
    super("w:numId");
    this.root.push({ _attr: { "w:val": typeof id === "string" ? `{${id}}` : id } });
  }
}

/**
 * Numbering change tracking (CT_TrackChangeNumbering).
 *
 * Records a revision to numbering properties.
 *
 * @internal
 */
class NumberingChange extends XmlComponent {
  public constructor(options: { original: string; id: string; author: string; date?: string }) {
    super("w:numberingChange");
    const attrs: Record<string, string> = {
      "w:original": options.original,
      "w:id": options.id,
      "w:author": options.author,
    };
    if (options.date !== undefined) attrs["w:date"] = options.date;
    this.root.push({ _attr: attrs });
  }
}
