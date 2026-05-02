/**
 * XML-serializable object types for OOXML document generation.
 *
 * @module
 */

/**
 * Attributes for an XML element.
 */
export type IXmlAttribute = Readonly<Record<string, string | number | boolean>>;

/**
 * Object that can be serialized to XML.
 */
export type IXmlableObject = Readonly<Record<string, any>>;
