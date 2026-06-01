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
 *
 * Uses `unknown` values (instead of `any`) for type safety at call sites,
 * while remaining compatible with the dynamic XML serialization pipeline.
 */
export type IXmlableObject = Readonly<Record<string, unknown>>;
