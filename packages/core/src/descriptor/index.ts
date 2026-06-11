/**
 * Descriptor system — declarative XML mapping for OOXML elements.
 *
 * @module
 */

// Types
export type {
  AttrSpec,
  ChildrenSpec,
  ChildSpec,
  ContentSpec,
  CustomDescriptor,
  ElementDescriptor,
  Descriptor,
  TextSpec,
  UnionSpec,
  UnionVariant,
} from "./types";

// Builder
export { DescriptorBuilder, element } from "./builder";

// Runtime
export { parse, stringify } from "./runtime";

// Context
export type { ReadContext, WriteContext } from "./context";

// OOXML Helpers
export { boolDecode, boolEncode, enumDecode, enumEncode } from "./helpers";

// Registry
export { DescriptorRegistry } from "./registry";
