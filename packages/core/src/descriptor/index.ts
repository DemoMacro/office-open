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

// Field consistency auditing — declared field sets + round-trip drift diff.
export { checkOrder, diffTagSets, roundTripFields } from "./field-consistency";
export type { FieldConsistencyReport, OrderViolation, RoundTripResult } from "./field-consistency";
export { FIELD_SPECS, findFieldSpec } from "./field-spec";
export type { DescriptorFieldSpec } from "./field-spec";
