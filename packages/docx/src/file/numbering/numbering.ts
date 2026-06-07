/**
 * Numbering module for WordprocessingML documents.
 *
 * Numbering provides support for numbered and bulleted lists.
 *
 * Reference: http://officeopenxml.com/WPnumbering.php
 *
 * @see https://stackoverflow.com/questions/58622437/purpose-of-abstractnum-and-numberinginstance
 *
 * @module
 */
import { AlignmentType } from "@file/paragraph";
import { XmlComponent, numberValObj } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import {
  abstractNumUniqueNumericIdGen,
  concreteNumUniqueNumericIdGen,
  convertInchesToTwip,
} from "@util/convenience-functions";
import { decimalNumber } from "@util/values";

import { buildDocumentAttributes } from "../document/document-attributes";
import { AbstractNumbering } from "./abstract-numbering";
import { LevelFormat } from "./level";
import type { LevelsOptions } from "./level";
import { ConcreteNumbering } from "./num";

/**
 * Options for configuring numbering definitions.
 *
 * @property config - Array of numbering configurations
 *
 * @see {@link Numbering}
 */
export interface NumberingOptions {
  /** Array of numbering configurations, each with levels and a reference name. */
  config: {
    /** Array of level definitions for this numbering configuration. */
    levels: LevelsOptions[];
    /** Unique reference name for this numbering configuration. */
    reference: string;
  }[];
  /** Numbering cleanup ID (w:numIdMacAtCleanup) */
  numIdMacAtCleanup?: number;
  /** Picture bullet definitions for numbering (w:numPicBullet) */
  numPicBullets?: {
    /** Unique ID for this picture bullet */
    numPicBulletId: number;
    /** Pict content as inline XML or reference */
    pict?: string;
  }[];
}

/**
 * Represents the numbering definitions in a WordprocessingML document.
 *
 * The numbering element contains abstract numbering definitions and their
 * concrete instances, which are referenced by paragraphs to create lists.
 * Each numbering configuration includes a default bullet list and any
 * custom numbering schemes defined by the user.
 *
 * Reference: http://officeopenxml.com/WPnumbering.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="numbering" type="CT_Numbering"/>
 *
 * <xsd:complexType name="CT_Numbering">
 *   <xsd:sequence>
 *     <xsd:element name="numPicBullet" type="CT_NumPicBullet" minOccurs="0" maxOccurs="unbounded"/>
 *     <xsd:element name="abstractNum" type="CT_AbstractNum" minOccurs="0" maxOccurs="unbounded"/>
 *     <xsd:element name="num" type="CT_Num" minOccurs="0" maxOccurs="unbounded"/>
 *     <xsd:element name="numIdMacAtCleanup" type="CT_DecimalNumber" minOccurs="0"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Create numbering with custom decimal list
 * const numbering = new Numbering({
 *   config: [
 *     {
 *       reference: "my-decimal-list",
 *       levels: [
 *         {
 *           level: 0,
 *           format: LevelFormat.DECIMAL,
 *           text: "%1.",
 *           alignment: AlignmentType.LEFT,
 *           start: 1,
 *           style: {
 *             paragraph: {
 *               indent: { left: 720, hanging: 360 },
 *             },
 *           },
 *         },
 *         {
 *           level: 1,
 *           format: LevelFormat.LOWER_LETTER,
 *           text: "%2)",
 *           alignment: AlignmentType.LEFT,
 *           style: {
 *             paragraph: {
 *               indent: { left: 1440, hanging: 360 },
 *             },
 *           },
 *         },
 *       ],
 *     },
 *   ],
 * });
 * ```
 */
export class Numbering extends XmlComponent {
  private abstractNumberingMap = new Map<string, AbstractNumbering>();
  private concreteNumberingMap = new Map<string, ConcreteNumbering>();
  private referenceConfigMap = new Map<string, Record<string, any>>();
  private abstractNumUniqueNumericId = abstractNumUniqueNumericIdGen();
  private concreteNumUniqueNumericId = concreteNumUniqueNumericIdGen();
  private _numIdMacAtCleanup?: number;
  private _numPicBullets?: {
    numPicBulletId: number;
    pict?: string;
  }[];

  /**
   * Creates a new numbering definition collection.
   *
   * Initializes the numbering with a default bullet list configuration and
   * any custom numbering configurations provided in the options.
   *
   * @param options - Configuration options for numbering definitions
   */
  public constructor(options: NumberingOptions) {
    super("w:numbering");
    this._numIdMacAtCleanup = options.numIdMacAtCleanup;
    this._numPicBullets = options.numPicBullets;
    this.root.push(
      buildDocumentAttributes(
        [
          "wpc",
          "mc",
          "o",
          "r",
          "m",
          "v",
          "wp14",
          "wp",
          "w10",
          "w",
          "w14",
          "w15",
          "wpg",
          "wpi",
          "wne",
          "wps",
        ],
        "w14 w15 wp14",
      ),
    );

    const abstractNumbering = new AbstractNumbering(this.abstractNumUniqueNumericId(), [
      {
        alignment: AlignmentType.LEFT,
        format: LevelFormat.BULLET,
        level: 0,
        style: {
          paragraph: {
            indent: {
              hanging: convertInchesToTwip(0.25),
              left: convertInchesToTwip(0.5),
            },
          },
        },
        text: "\u25CF",
      },
      {
        alignment: AlignmentType.LEFT,
        format: LevelFormat.BULLET,
        level: 1,
        style: {
          paragraph: {
            indent: {
              hanging: convertInchesToTwip(0.25),
              left: convertInchesToTwip(1),
            },
          },
        },
        text: "\u25CB",
      },
      {
        alignment: AlignmentType.LEFT,
        format: LevelFormat.BULLET,
        level: 2,
        style: {
          paragraph: {
            indent: { hanging: convertInchesToTwip(0.25), left: 2160 },
          },
        },
        text: "\u25A0",
      },
      {
        alignment: AlignmentType.LEFT,
        format: LevelFormat.BULLET,
        level: 3,
        style: {
          paragraph: {
            indent: { hanging: convertInchesToTwip(0.25), left: 2880 },
          },
        },
        text: "\u25CF",
      },
      {
        alignment: AlignmentType.LEFT,
        format: LevelFormat.BULLET,
        level: 4,
        style: {
          paragraph: {
            indent: { hanging: convertInchesToTwip(0.25), left: 3600 },
          },
        },
        text: "\u25CB",
      },
      {
        alignment: AlignmentType.LEFT,
        format: LevelFormat.BULLET,
        level: 5,
        style: {
          paragraph: {
            indent: { hanging: convertInchesToTwip(0.25), left: 4320 },
          },
        },
        text: "\u25A0",
      },
      {
        alignment: AlignmentType.LEFT,
        format: LevelFormat.BULLET,
        level: 6,
        style: {
          paragraph: {
            indent: { hanging: convertInchesToTwip(0.25), left: 5040 },
          },
        },
        text: "\u25CF",
      },
      {
        alignment: AlignmentType.LEFT,
        format: LevelFormat.BULLET,
        level: 7,
        style: {
          paragraph: {
            indent: { hanging: convertInchesToTwip(0.25), left: 5760 },
          },
        },
        text: "\u25CF",
      },
      {
        alignment: AlignmentType.LEFT,
        format: LevelFormat.BULLET,
        level: 8,
        style: {
          paragraph: {
            indent: { hanging: convertInchesToTwip(0.25), left: 6480 },
          },
        },
        text: "\u25CF",
      },
    ]);

    this.concreteNumberingMap.set(
      "default-bullet-numbering",
      new ConcreteNumbering({
        abstractNumId: abstractNumbering.id,
        instance: 0,
        numId: 1,
        overrideLevels: [
          {
            num: 0,
            start: 1,
          },
        ],
        reference: "default-bullet-numbering",
      }),
    );

    this.abstractNumberingMap.set("default-bullet-numbering", abstractNumbering);

    for (const con of options.config) {
      this.abstractNumberingMap.set(
        con.reference,
        new AbstractNumbering(this.abstractNumUniqueNumericId(), con.levels),
      );
      this.referenceConfigMap.set(con.reference, con.levels);
    }
  }

  public override toXml(context: Context): string {
    const start = this.root.length;

    // numPicBullet elements come first (XSD order)
    if (this._numPicBullets) {
      for (const bullet of this._numPicBullets) {
        this.root.push(new NumPicBullet(bullet.numPicBulletId, bullet.pict));
      }
    }

    for (const numbering of this.abstractNumberingMap.values()) {
      this.root.push(numbering);
    }
    for (const numbering of this.concreteNumberingMap.values()) {
      this.root.push(numbering);
    }
    if (this._numIdMacAtCleanup !== undefined) {
      this.root.push(numberValObj("w:numIdMacAtCleanup", decimalNumber(this._numIdMacAtCleanup)));
    }
    const result = super.toXml(context);
    this.root.length = start;
    return result;
  }

  /**
   * Creates a concrete numbering instance from an abstract numbering definition.
   *
   * This method creates a new concrete numbering instance that references an
   * abstract numbering definition. It's used internally when paragraphs reference
   * numbering configurations.
   *
   * @param reference - The reference name of the abstract numbering definition
   * @param instance - The instance number for this concrete numbering
   */
  public createConcreteNumberingInstance(reference: string, instance: number): void {
    const abstractNumbering = this.abstractNumberingMap.get(reference);

    if (!abstractNumbering) {
      return;
    }

    const fullReference = `${reference}-${instance}`;

    if (this.concreteNumberingMap.has(fullReference)) {
      return;
    }

    const referenceConfigLevels = this.referenceConfigMap.get(reference);
    const firstLevelStartNumber = referenceConfigLevels && referenceConfigLevels[0].start;

    const concreteNumberingSettings = {
      abstractNumId: abstractNumbering.id,
      instance,
      numId: this.concreteNumUniqueNumericId(),
      overrideLevels: [
        typeof firstLevelStartNumber === "number" && Number.isInteger(firstLevelStartNumber)
          ? {
              num: 0,
              start: firstLevelStartNumber,
            }
          : {
              num: 0,
              start: 1,
            },
      ],
      reference,
    };

    this.concreteNumberingMap.set(fullReference, new ConcreteNumbering(concreteNumberingSettings));
  }

  /**
   * Gets all concrete numbering instances.
   *
   * @returns An array of all concrete numbering instances
   */
  public get concreteNumbering(): ConcreteNumbering[] {
    return [...this.concreteNumberingMap.values()];
  }

  /**
   * Gets all reference configurations.
   *
   * @returns An array of all numbering reference configurations
   */
  public get referenceConfig(): Record<string, any>[] {
    return [...this.referenceConfigMap.values()];
  }
}

/**
 * Picture bullet definition (CT_NumPicBullet).
 *
 * Defines a picture to be used as a bullet in a numbering definition.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_NumPicBullet
 */
class NumPicBullet extends XmlComponent {
  public constructor(id: number, pict?: string) {
    super("w:numPicBullet");
    this.root.push(numberValObj("w:numPicBulletId", id));
    if (pict) {
      // pict is a raw XML string for w:pict content
      this.root.push({ _raw: pict } as any);
    }
  }
}
