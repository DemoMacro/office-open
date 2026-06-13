/**
 * Descriptor system — bidirectional XML mapping for OOXML parts.
 *
 * @module
 */

// Types
export type { CustomDescriptor, Descriptor } from "./types";

// Runtime
export { parse, stringify } from "./runtime";

// Context
export type { ReadContext, WriteContext } from "./context";

// OOXML Helpers
export { boolDecode, boolEncode, enumDecode, enumEncode } from "./helpers";

// Registry
export { DescriptorRegistry } from "./registry";
