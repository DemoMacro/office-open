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
  verticalMerge?: (typeof VerticalMergeRevisionType)[keyof typeof VerticalMergeRevisionType];
  verticalMergeOriginal?: (typeof VerticalMergeRevisionType)[keyof typeof VerticalMergeRevisionType];
};
