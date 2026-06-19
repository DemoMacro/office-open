/**
 * Comment types for WordprocessingML documents.
 *
 * @module
 */

import type { ParagraphOptions } from "@parts/paragraph/paragraph";

/**
 * Options for creating a single comment.
 */
export interface CommentOptions {
  /** Unique identifier for the comment */
  id: number;
  /** Content of the comment (paragraphs) */
  children: (string | ParagraphOptions)[];

  /** Initials of the comment author */
  initials?: string;
  /** Name of the comment author */
  author?: string;
  /** Date and time the comment was created */
  date?: Date | string;
}

/**
 * Options for creating a comments container.
 */
export interface CommentsOptions {
  /** Array of comment definitions */
  children: CommentOptions[];
}
