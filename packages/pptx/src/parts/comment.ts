/**
 * Comment types for PPTX.
 *
 * @module
 */
import type { UniversalMeasure } from "@office-open/core";

export interface AuthorEntry {
  id: number;
  name: string;
  initials: string;
  clrIdx: number;
  lastIdx: number;
}

export interface CommentEntry {
  authorId: number;
  idx: number;
  date?: string;
  modified?: boolean;
  x: number | UniversalMeasure;
  y: number | UniversalMeasure;
  text: string;
}
