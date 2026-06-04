/**
 * ProofError module for WordprocessingML documents.
 *
 * Proof errors mark the start and end of spelling and grammar error ranges.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_ProofErr
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

/**
 * Proof error type constants.
 *
 * These values specify whether the proof error marks the start or end
 * of a spelling or grammar error range.
 */
export const ProofErrorType = {
  /** Start of a spelling error range */
  SPELL_START: "spellStart",
  /** End of a spelling error range */
  SPELL_END: "spellEnd",
  /** Start of a grammar error range */
  GRAM_START: "gramStart",
  /** End of a grammar error range */
  GRAM_END: "gramEnd",
} as const;

export type ProofErrorTypeValue = (typeof ProofErrorType)[keyof typeof ProofErrorType];

/**
 * Represents a proofing error marker (w:proofErr) in a WordprocessingML document.
 *
 * Proof errors delimit spelling and grammar error regions within paragraph content.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_ProofErr
 *
 * @example
 * ```typescript
 * // Mark a spelling error range
 * new ProofError(ProofErrorType.SPELL_START);
 * new TextRun("teh");
 * new ProofError(ProofErrorType.SPELL_END);
 * ```
 */
export class ProofError extends XmlComponent {
  public constructor(type: ProofErrorTypeValue) {
    super("w:proofErr");
    this.root.push({ _attr: { "w:type": type } });
  }
}
