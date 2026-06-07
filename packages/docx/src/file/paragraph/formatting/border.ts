/**
 * Paragraph border module for WordprocessingML documents.
 *
 * This module provides border options for paragraphs.
 *
 * Reference: http://officeopenxml.com/WPborders.php
 *
 * @module
 */
import { BorderStyle, buildBorderObj, createBorderElement } from "@file/border";
import type { BorderOptions } from "@file/border";
import type { IXmlableObject } from "@file/xml-components";
import { IgnoreIfEmptyXmlComponent, XmlComponent } from "@file/xml-components";

/**
 * Build paragraph borders (w:pBdr) as IXmlableObject without allocating XmlComponent tree.
 */
export function buildParagraphBorders(options: BordersOptions): IXmlableObject | undefined {
  const children: IXmlableObject[] = [];

  if (options.top) children.push(buildBorderObj("w:top", options.top));
  if (options.left) children.push(buildBorderObj("w:left", options.left));
  if (options.bottom) children.push(buildBorderObj("w:bottom", options.bottom));
  if (options.right) children.push(buildBorderObj("w:right", options.right));
  if (options.between) children.push(buildBorderObj("w:between", options.between));
  if (options.bar) children.push(buildBorderObj("w:bar", options.bar));

  return children.length > 0 ? { "w:pBdr": children } : undefined;
}

/**
 * Build thematic break (w:pBdr with bottom border) as IXmlableObject.
 */
export function buildThematicBreakObj(): IXmlableObject {
  return {
    "w:pBdr": [
      buildBorderObj("w:bottom", {
        color: "auto",
        size: 6,
        space: 1,
        style: BorderStyle.SINGLE,
      }),
    ],
  };
}

/**
 * Options for configuring paragraph borders.
 *
 * Borders can be applied to top, bottom, left, right, and between paragraphs.
 *
 * @property top - Border for the top edge of the paragraph
 * @property bottom - Border for the bottom edge of the paragraph
 * @property left - Border for the left edge of the paragraph
 * @property right - Border for the right edge of the paragraph
 * @property between - Border between consecutive paragraphs with the same border settings
 */
export interface BordersOptions {
  /** Border for the top edge of the paragraph */
  top?: BorderOptions;
  /** Border for the bottom edge of the paragraph */
  bottom?: BorderOptions;
  /** Border for the left edge of the paragraph */
  left?: BorderOptions;
  /** Border for the right edge of the paragraph */
  right?: BorderOptions;
  /** Border between consecutive paragraphs with the same border settings */
  between?: BorderOptions;
  /** Bar border (paragraph-level bar, rendered at left of paragraph) */
  bar?: BorderOptions;
}

/**
 * Represents paragraph borders in a WordprocessingML document.
 *
 * The pBdr element specifies borders that surround the paragraph.
 *
 * Reference: http://officeopenxml.com/WPborders.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_PBdr">
 *   <xsd:sequence>
 *     <xsd:element name="top" type="CT_Border" minOccurs="0"/>
 *     <xsd:element name="left" type="CT_Border" minOccurs="0"/>
 *     <xsd:element name="bottom" type="CT_Border" minOccurs="0"/>
 *     <xsd:element name="right" type="CT_Border" minOccurs="0"/>
 *     <xsd:element name="between" type="CT_Border" minOccurs="0"/>
 *     <xsd:element name="bar" type="CT_Border" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * new Paragraph({
 *   border: {
 *     top: { style: BorderStyle.SINGLE, size: 6, color: "FF0000" },
 *     bottom: { style: BorderStyle.SINGLE, size: 6, color: "FF0000" },
 *   },
 *   children: [new TextRun("Paragraph with top and bottom borders")],
 * });
 * ```
 */
export class Border extends IgnoreIfEmptyXmlComponent {
  public constructor(options: BordersOptions) {
    super("w:pBdr");

    if (options.top) {
      this.root.push(createBorderElement("w:top", options.top));
    }

    if (options.left) {
      this.root.push(createBorderElement("w:left", options.left));
    }

    if (options.bottom) {
      this.root.push(createBorderElement("w:bottom", options.bottom));
    }

    if (options.right) {
      this.root.push(createBorderElement("w:right", options.right));
    }

    if (options.between) {
      this.root.push(createBorderElement("w:between", options.between));
    }

    if (options.bar) {
      this.root.push(createBorderElement("w:bar", options.bar));
    }
  }
}

/**
 * Represents a thematic break (horizontal rule) in a WordprocessingML document.
 *
 * Creates a horizontal line across the paragraph using a bottom border.
 *
 * Reference: http://officeopenxml.com/WPborders.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_PBdr">
 *   <xsd:sequence>
 *     <xsd:element name="bottom" type="CT_Border" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * new Paragraph({
 *   thematicBreak: true,
 *   children: [new TextRun("Paragraph with horizontal rule below")],
 * });
 * ```
 */
export class ThematicBreak extends XmlComponent {
  public constructor() {
    super("w:pBdr");
    const bottom = createBorderElement("w:bottom", {
      color: "auto",
      size: 6,
      space: 1,
      style: BorderStyle.SINGLE,
    });
    this.root.push(bottom);
  }
}
