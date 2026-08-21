/**
 * Positional tab types for WordprocessingML documents.
 *
 * @module
 */

export const PositionalTabAlignment = {
  LEFT: "left",
  CENTER: "center",
  RIGHT: "right",
} as const;

export const PositionalTabRelativeTo = {
  MARGIN: "margin",
  INDENT: "indent",
} as const;

export const PositionalTabLeader = {
  NONE: "none",
  DOT: "dot",
  HYPHEN: "hyphen",
  UNDERSCORE: "underscore",
  MIDDLE_DOT: "middleDot",
} as const;

export interface PositionalTabOptions {
  /** Where the tab lands on the reference line. */
  alignment: (typeof PositionalTabAlignment)[keyof typeof PositionalTabAlignment];
  /** Reference line for the position: page margin or paragraph indent. */
  relativeTo: (typeof PositionalTabRelativeTo)[keyof typeof PositionalTabRelativeTo];
  /** Fill character drawn across the tab. */
  leader: (typeof PositionalTabLeader)[keyof typeof PositionalTabLeader];
}
