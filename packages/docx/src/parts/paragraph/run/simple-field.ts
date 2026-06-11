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
}
