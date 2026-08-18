/**
 * Comment types for WordprocessingML documents.
 *
 * @module
 */

import type { BookmarkStartOptions, MarkupRangeOptions } from "@parts/paragraph/links/bookmark";
import type { ParagraphOptions } from "@parts/paragraph/paragraph";
import type { TableOptions } from "@parts/table/table";

import type { RunOptions } from "./run";

/**
 * Options for creating a single comment.
 */
export interface CommentOptions {
  /** Unique identifier for the comment */
  id: number;
  /** Block-level comment content (paragraphs, tables) plus comment-level
   *  bookmark markers Word anchors directly inside w:comment. */
  children: (
    | string
    | ParagraphOptions
    | { table: TableOptions }
    | { bookmarkStart: BookmarkStartOptions }
    | { bookmarkEnd: MarkupRangeOptions }
  )[];

  /** Initials of the comment author */
  initials?: string;
  /** Name of the comment author */
  author?: string;
  /** Creation date (CT_Comment `@w:date`, ISO 8601 string) */
  date?: string;
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
  /** Comment author (CT_Comment `@w:author` — required by XSD, defaults to ""). */
  author?: string;
  /** Author initials (CT_Comment `@w:initials`). */
  initials?: string;
  /** Creation date (CT_Comment `@w:date`, ISO 8601 string); defaults to the current time. */
  date?: string;
  /** Comment reply content stored in word/comments.xml (maps to CommentOptions.children). */
  children: (string | ParagraphOptions)[];
  /** Anchored document content the comment range wraps (inline runs/text). */
  wrap?: (string | RunOptions)[];
}
