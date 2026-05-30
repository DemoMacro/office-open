/**
 * XML attribute types for OOXML document generation.
 *
 * @module
 */

/**
 * Structured attribute payload with explicit key-value mapping.
 */
export type AttributePayload<T> = {
  readonly [P in keyof T]: { readonly key: string; readonly value: T[P] };
};
