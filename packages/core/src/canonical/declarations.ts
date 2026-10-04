/**
 * A named sequence-field declaration.
 */
export interface SequenceDeclarationOptions {
  /** Unique sequence name used by sequence fields. */
  name: string;
  /** Outline level shown in the sequence number; omit for a flat number. */
  displayOutlineLevel?: number;
  /** Text between the outline level and sequence number. */
  outlineSeparator?: string;
}

/**
 * A named variable-field declaration.
 */
export interface VariableDeclarationOptions {
  /** Unique variable name used by variable fields. */
  name: string;
  /** Stored value type; display text is owned by the field result. */
  valueType: "float" | "string";
}
