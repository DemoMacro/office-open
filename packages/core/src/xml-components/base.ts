/**
 * Base XML Component infrastructure for OOXML document generation.
 *
 * @module
 */
import { xml } from "@office-open/xml";

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
  readonly stack: IXmlableObject[];
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
   * Prepares this component for XML serialization.
   * Override in subclasses that build IXmlableObject trees.
   * Classes overriding toXml() for direct string output do not need to override this.
   *
   * @param context - The serialization context
   * @returns The XML-serializable object, or undefined to exclude from output
   */
  public prepForXml(context: Context): IXmlableObject | undefined {
    return undefined;
  }

  /**
   * Direct XML serialization. Override in subclasses for zero-allocation output.
   * Default falls back to prepForXml() + xml().
   */
  public toXml(context: Context): string {
    const obj = this.prepForXml(context);
    return obj ? xml(obj) : "";
  }
}
