/**
 * Comment types for PPTX.
 *
 * @module
 */

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
  x: number;
  y: number;
  text: string;
}
