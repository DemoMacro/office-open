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
