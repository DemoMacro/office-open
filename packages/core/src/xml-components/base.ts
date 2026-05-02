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
export interface IContext<TFileData = unknown> {
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
     *
     * @param context - The serialization context
     * @returns The XML-serializable object, or undefined to exclude from output
     */
    public abstract prepForXml(context: IContext): IXmlableObject | undefined;
}
