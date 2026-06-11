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
  alignment: (typeof PositionalTabAlignment)[keyof typeof PositionalTabAlignment];
  relativeTo: (typeof PositionalTabRelativeTo)[keyof typeof PositionalTabRelativeTo];
  leader: (typeof PositionalTabLeader)[keyof typeof PositionalTabLeader];
}
