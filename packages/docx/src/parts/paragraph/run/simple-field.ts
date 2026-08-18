/**
 * Simple field types for WordprocessingML documents.
 *
 * @module
 */

export interface SimpleFieldOptions {
  /** Field instruction string */
  instruction: string;
  /** Optional cached field value */
  cachedValue?: string;
  /** Verbatim XML of the cached-value runs when they carry run properties
   *  (Word marks field results with w:noProof/rFonts) or split across runs —
   *  shapes the plain single-text-run template cannot reproduce. */
  cachedRunsXml?: string;
  /** Lock the field to prevent updates (CT_SimpleField `@fldLock`) */
  fieldLock?: boolean;
  /** Field result is out of date (CT_SimpleField `@dirty`) */
  dirty?: boolean;
}
