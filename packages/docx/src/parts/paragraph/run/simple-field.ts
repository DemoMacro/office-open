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
  /** Lock the field to prevent updates (CT_SimpleField @fldLock) */
  fldLock?: boolean;
  /** Field result is out of date (CT_SimpleField @dirty) */
  dirty?: boolean;
}
