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
  /** Cached-value runs when the plain single-text-run template cannot express
   *  their run properties, identity attributes, or child sequence. */
  cachedRuns?: ParagraphChild[];
  /** Lock the field to prevent updates (CT_SimpleField `@fldLock`) */
  fieldLock?: boolean;
  /** Field result is out of date (CT_SimpleField `@dirty`) */
  dirty?: boolean;
}

import type { ParagraphChild } from "../paragraph";
