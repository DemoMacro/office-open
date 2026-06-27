/**
 * Comment types for WordprocessingML documents.
 *
 * @module
 */

import type { ParagraphOptions } from "@parts/paragraph/paragraph";

import type { RunOptions } from "./run";

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

/**
 * A comment authored as a single inline paragraph child. The library allocates
 * the comment id, emits the range markers + reference, and registers the comment
 * entry in word/comments.xml — the caller never touches an id or pairs markers.
 *
 * `children` is the comment reply (stored in the comments part, same shape as
 * {@link CommentOptions.children}); `wrap` is the anchored document content the
 * comment range wraps (inline runs/text, emitted between the range markers).
 *
 * Reference: wml.xsd CT_Markup, CT_Comment, EG_RangeMarkupElements.
 */
export interface CommentChildOptions {
  /** Comment author (CT_Comment @w:author — required by XSD, defaults to ""). */
  author?: string;
  /** Author initials (CT_Comment @w:initials). */
  initials?: string;
  /** Creation date (CT_Comment @w:date); defaults to the current time. */
  date?: Date | string;
  /** Comment reply content stored in word/comments.xml (maps to CommentOptions.children). */
  children: (string | ParagraphOptions)[];
  /** Anchored document content the comment range wraps (inline runs/text). */
  wrap?: (string | RunOptions)[];
}
