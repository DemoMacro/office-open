/**
 * Cell merge track revision component.
 *
 * @module
 */

import type { ChangedProperties } from "../track-revision";

/**
 * Vertical merge revision types.
 */
export const VerticalMergeRevisionType = {
  /**
   * Cell that is merged with upper one.
   */
  CONTINUE: "continue",
  /**
   * Cell that is starting the vertical merge.
   */
  RESTART: "restart",
} as const;

export type CellMergeAttributes = ChangedProperties & {
  /** Merge state this revision sets: "restart" first cell of the merge, "continue" cell absorbed into it. */
  verticalMerge?: (typeof VerticalMergeRevisionType)[keyof typeof VerticalMergeRevisionType];
  /** Merge state before the revision. */
  verticalMergeOriginal?: (typeof VerticalMergeRevisionType)[keyof typeof VerticalMergeRevisionType];
};
