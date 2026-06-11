/**
 * Permission range markers for WordprocessingML document protection.
 *
 * These elements mark editable ranges within a protected document.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_PermStart, CT_PermEnd
 *
 * @module
 */

/**
 * Editing group values for permission ranges.
 *
 * Defines who can edit within a permission range.
 *
 * Reference: ISO/IEC 29500-4, ST_EdGrp
 */
export const EditGroupType = {
  NONE: "none",
  EVERYONE: "everyone",
  ADMINISTRATORS: "administrators",
  CONTRIBUTORS: "contributors",
  EDITORS: "editors",
  OWNERS: "owners",
  CURRENT: "current",
} as const;

export type EditGroup = (typeof EditGroupType)[keyof typeof EditGroupType];

/**
 * Options for creating a permission start marker.
 */
export interface PermStartOptions {
  /** Unique identifier for this permission range (typically a number) */
  id: string | number;
  /** Editing group that can edit this range */
  edGroup?: EditGroup;
  /** Individual user who can edit this range */
  ed?: string;
  /** First column this range covers (for table cells) */
  colFirst?: number;
  /** Last column this range covers (for table cells) */
  colLast?: number;
}
