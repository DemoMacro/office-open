/**
 * Abstract numbering definitions module for WordprocessingML documents.
 *
 * Abstract numbering definitions contain the formatting and style information
 * that can be shared across multiple numbering instances.
 *
 * Reference: http://officeopenxml.com/WPnumbering.php
 *
 * @module
 */
import { XmlComponent, stringValObj } from "@file/xml-components";
import { decimalNumber } from "@util/values";

import { Level } from "./level";
import type { LevelsOptions } from "./level";
import { MultiLevelType } from "./multi-level-type";

/**
 * Options for abstract numbering definitions.
 */
export interface AbstractNumberingOptions {
  /** Unique hex identifier for this numbering set */
  readonly nsid?: string;
  /** Template hex identifier */
  readonly tmpl?: string;
  /** Name of the numbering definition */
  readonly name?: string;
  /** Paragraph style that references this numbering */
  readonly styleLink?: string;
  /** Numbering style to inherit from */
  readonly numStyleLink?: string;
}

/**
 * Represents an abstract numbering definition in a WordprocessingML document.
 *
 * Reference: http://officeopenxml.com/WPnumbering.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_AbstractNum">
 *   <xsd:sequence>
 *     <xsd:element name="nsid" type="CT_LongHexNumber" minOccurs="0"/>
 *     <xsd:element name="multiLevelType" type="CT_MultiLevelType" minOccurs="0"/>
 *     <xsd:element name="tmpl" type="CT_LongHexNumber" minOccurs="0"/>
 *     <xsd:element name="name" type="CT_String" minOccurs="0"/>
 *     <xsd:element name="styleLink" type="CT_String" minOccurs="0"/>
 *     <xsd:element name="numStyleLink" type="CT_String" minOccurs="0"/>
 *     <xsd:element name="lvl" type="CT_Lvl" minOccurs="0" maxOccurs="9"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="abstractNumId" type="ST_DecimalNumber" use="required"/>
 * </xsd:complexType>
 * ```
 */
export class AbstractNumbering extends XmlComponent {
  /** The unique identifier for this abstract numbering definition. */
  public readonly id: number;

  /**
   * Creates a new abstract numbering definition.
   */
  public constructor(
    id: number,
    levelOptions: readonly LevelsOptions[],
    extraOptions?: AbstractNumberingOptions,
  ) {
    super("w:abstractNum");
    this.root.push({
      _attr: {
        "w:abstractNumId": decimalNumber(id),
        "w15:restartNumberingAfterBreak": 0,
      },
    });
    this.id = id;

    if (extraOptions?.nsid !== undefined) {
      this.root.push(stringValObj("w:nsid", extraOptions.nsid));
    }

    this.root.push(new MultiLevelType("hybridMultilevel"));

    if (extraOptions?.tmpl !== undefined) {
      this.root.push(stringValObj("w:tmpl", extraOptions.tmpl));
    }
    if (extraOptions?.name !== undefined) {
      this.root.push(stringValObj("w:name", extraOptions.name));
    }
    if (extraOptions?.styleLink !== undefined) {
      this.root.push(stringValObj("w:styleLink", extraOptions.styleLink));
    }
    if (extraOptions?.numStyleLink !== undefined) {
      this.root.push(stringValObj("w:numStyleLink", extraOptions.numStyleLink));
    }

    for (const option of levelOptions) {
      this.root.push(new Level(option));
    }
  }
}
