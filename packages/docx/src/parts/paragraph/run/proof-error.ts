/**
 * Proof error types for WordprocessingML documents.
 *
 * @module
 */

export const ProofErrorType = {
  SPELL_START: "spellStart",
  SPELL_END: "spellEnd",
  GRAM_START: "gramStart",
  GRAM_END: "gramEnd",
} as const;

export type ProofErrorTypeValue = (typeof ProofErrorType)[keyof typeof ProofErrorType];
