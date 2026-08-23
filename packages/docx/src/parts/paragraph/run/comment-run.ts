/**
 * Comment types for WordprocessingML documents.
 *
 * @module
 */

import type { DateTime } from "@office-open/core";
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
  /**
   * Creation date (CT_Comment `@w:date`, ISO 8601 string). Tri-stated:
   * `null` = the source comment carried no date (attribute omitted),
   * omitted = fresh authoring (defaults to the current time), a string = used
   * verbatim.
   */
  date?: string | null;
}

/**
 * Inline comment sugar: the library allocates the id, emits the range markers +
 * reference, and registers the entry in word/comments.xml. `children` is the
 * comment reply (same shape as CommentOptions.children); `wrap` is the anchored
 * content the comment range wraps.
 */
export interface CommentChildOptions {
  /** Comment author (CT_Comment `@w:author` — required by XSD, defaults to ""). */
  author?: string;
  /** Author initials (CT_Comment `@w:initials`). */
  initials?: string;
  /** Creation date (CT_Comment `@w:date`, ISO 8601 string); defaults to the current time. */
  date?: DateTime;
  /** Comment reply content stored in word/comments.xml (maps to CommentOptions.children). */
  children: (string | ParagraphOptions)[];
  /** Anchored document content the comment range wraps (inline runs/text). */
  wrap?: (string | RunOptions)[];
}
