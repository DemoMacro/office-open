/**
 * Break type values for WordprocessingML documents.
 *
 * @module
 */

/**
 * Break type constants.
 * @internal
 */
export const BreakType = {
  /** Column break */
  COLUMN: "column",
  /** Page break */
  PAGE: "page",
} as const;

export type BreakTypeValue = (typeof BreakType)[keyof typeof BreakType];
