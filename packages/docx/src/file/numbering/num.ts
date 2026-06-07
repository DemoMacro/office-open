/**
 * Concrete numbering instances module for WordprocessingML documents.
 *
 * Concrete numbering instances reference abstract numbering definitions and
 * can override specific level settings. Each paragraph references a concrete
 * numbering instance to apply list formatting.
 *
 * Reference: http://officeopenxml.com/WPnumbering.php
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";
import { decimalNumber } from "@util/values";

/**
 * Reference to an abstract numbering definition.
 */
class AbstractNumId extends XmlComponent {
  public constructor(value: number) {
    super("w:abstractNumId");
    this.root.push({ _attr: { "w:val": value } });
  }
}

/**
 * Options for overriding a specific level in a numbering instance.
 *
 * @property num - The level number to override (0-8)
 * @property start - The starting number for this level
 */
interface OverrideLevel {
  /** The level number to override (0-8). */
  num: number;
  /** The starting number for this level. */
  start?: number;
}

/**
 * Options for creating a concrete numbering instance.
 *
 * @property numId - Unique identifier for this numbering instance
 * @property abstractNumId - ID of the abstract numbering definition to reference
 * @property reference - Reference name for this numbering instance
 * @property instance - Instance number for tracking multiple uses
 * @property overrideLevels - Array of level overrides to customize specific levels
 */
export interface ConcreteNumberingOptions {
  /** Unique identifier for this numbering instance. */
  numId: number;
  /** ID of the abstract numbering definition to reference. */
  abstractNumId: number;
  /** Reference name for this numbering instance. */
  reference: string;
  /** Instance number for tracking multiple uses. */
  instance: number;
  /** Array of level overrides to customize specific levels. */
  overrideLevels?: OverrideLevel[];
}

/**
 * Represents a concrete numbering instance in a WordprocessingML document.
 *
 * A concrete numbering instance references an abstract numbering definition and
 * can override specific levels. Paragraphs reference concrete numbering instances
 * to apply list formatting.
 *
 * Reference: http://officeopenxml.com/WPnumbering.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Num">
 *   <xsd:sequence>
 *     <xsd:element name="abstractNumId" type="CT_DecimalNumber" minOccurs="1"/>
 *     <xsd:element name="lvlOverride" type="CT_NumLvl" minOccurs="0" maxOccurs="9"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="numId" type="ST_DecimalNumber" use="required"/>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Create a concrete numbering instance
 * const concreteNumbering = new ConcreteNumbering({
 *   numId: 1,
 *   abstractNumId: 0,
 *   reference: "my-numbering",
 *   instance: 0,
 *   overrideLevels: [
 *     {
 *       num: 0,
 *       start: 5, // Start numbering at 5 instead of 1
 *     },
 *   ],
 * });
 * ```
 */
export class ConcreteNumbering extends XmlComponent {
  /** The unique identifier for this numbering instance. */
  public numId: number;
  /** The reference name for this numbering instance. */
  public reference: string;
  /** The instance number for tracking multiple uses. */
  public instance: number;

  /**
   * Creates a new concrete numbering instance.
   *
   * @param options - Configuration options for the numbering instance
   */
  public constructor(options: ConcreteNumberingOptions) {
    super("w:num");

    this.numId = options.numId;
    this.reference = options.reference;
    this.instance = options.instance;

    this.root.push({ _attr: { "w:numId": decimalNumber(options.numId) } });

    this.root.push(new AbstractNumId(decimalNumber(options.abstractNumId)));

    if (options.overrideLevels && options.overrideLevels.length) {
      for (const level of options.overrideLevels) {
        this.root.push(new LevelOverride(level.num, level.start));
      }
    }
  }
}

/**
 * Represents a level override in a concrete numbering instance.
 *
 * Level overrides allow customization of specific levels within a numbering
 * instance, such as changing the starting number.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_NumLvl">
 *   <xsd:sequence>
 *     <xsd:element name="startOverride" type="CT_DecimalNumber" minOccurs="0"/>
 *     <xsd:element name="lvl" type="CT_Lvl" minOccurs="0" maxOccurs="1"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="ilvl" type="ST_DecimalNumber" use="required"/>
 * </xsd:complexType>
 * ```
 */
export class LevelOverride extends XmlComponent {
  /**
   * Creates a new level override.
   *
   * @param levelNum - The level number to override (0-8)
   * @param start - Optional starting number for the level
   */
  public constructor(levelNum: number, start?: number) {
    super("w:lvlOverride");
    this.root.push({ _attr: { "w:ilvl": levelNum } });
    if (start !== undefined) {
      this.root.push(new StartOverride(start));
    }
  }
}

class StartOverride extends XmlComponent {
  public constructor(start: number) {
    super("w:startOverride");
    this.root.push({ _attr: { "w:val": start } });
  }
}
