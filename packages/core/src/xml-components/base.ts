/**
 * Base XML Component infrastructure for OOXML document generation.
 *
 * @module
 */
import type { IXmlableObject } from "./types";
export type { IXmlableObject };

/**
 * Context object passed through the XML tree during serialization.
 *
 * @typeParam TFileData - The type of the root file data object (format-specific)
 */
export interface Context<TFileData = unknown> {
  /** The root file data object being serialized (format-specific). */
  readonly fileData?: TFileData;
  /** Current traversal stack of components (mutable for performance). */
  readonly stack: BaseXmlComponent[];
}

/**
 * Abstract base class for all XML components.
 */
export abstract class BaseXmlComponent {
  /** The XML element name for this component (e.g., "w:p" for paragraph). */
  protected readonly rootKey: string;

  public constructor(rootKey: string) {
    this.rootKey = rootKey;
  }

  /**
   * Direct XML serialization. Override in subclasses for zero-allocation output.
   */
  public toXml(_context: Context): string {
    return "";
  }

  /**
   * Context-free XML serialization.
   *
   * For components that do not depend on `Context`, this provides a
   * clean alternative to `toXml({ stack: [] })`.
   * The default implementation passes a dummy context.
   */
  public serialize(): string {
    return this.toXml({ stack: [] });
  }
}
