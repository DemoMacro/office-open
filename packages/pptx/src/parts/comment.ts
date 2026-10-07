/**
 * Comment types for PPTX.
 *
 * @module
 */
import type { UniversalMeasure } from "@office-open/core";
import type { DateTime } from "@office-open/core";

export interface AuthorEntry {
  id: number;
  name: string;
  initials: string;
  clrIdx: number;
  lastIdx: number;
  /** Verbatim inner XML of p:extLst — round-trip only. */
  ext?: string;
}

export interface CommentEntry {
  authorId: number;
  idx: number;
  date?: DateTime;
  modified?: boolean;
  /** Verbatim inner XML of p:extLst — round-trip only. */
  ext?: string;
  x: number | UniversalMeasure;
  y: number | UniversalMeasure;
  text: string;
}
