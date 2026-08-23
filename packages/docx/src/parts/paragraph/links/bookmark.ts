/**
 * Bookmark types for WordprocessingML documents.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd — CT_Markup, CT_MarkupRange,
 * CT_BookmarkRange, CT_Bookmark.
 *
 * @module
 */

import type { DateTime } from "@office-open/core";
import type { RunOptions } from "@parts/paragraph/run/run";

/**
 * ST_DisplacedByCustomXml — whether a range marker is displaced to the
 * previous or next sibling customXml element. Shared by every
 * CT_MarkupRange derivative (bookmark, move range, comment range, …).
 */
export type DisplacedByCustomXml = "next" | "prev";

/**
 * Range end marker (`@id` + `@displacedByCustomXml`), shared by w:bookmarkEnd,
 * w:commentRangeStart/End, and w:moveFrom/ToRangeEnd. End markers carry no name
 * or column range (only start markers do).
 */
export interface MarkupRangeOptions {
  /** Marker identifier (CT_Markup `@w:id`, required). */
  id: number;
  /** Displacement relative to a sibling customXml (CT_MarkupRange `@w:displacedByCustomXml`). */
  displacedByCustomXml?: DisplacedByCustomXml;
}

/**
 * Options for a bookmark start (w:bookmarkStart).
 *
 * Maps to CT_Bookmark = CT_BookmarkRange (colFirst/colLast) + name. The column
 * range scopes a bookmark to specific table columns so Word preserves the exact
 * cell span on round-trip rather than recomputing it.
 */
export interface BookmarkStartOptions extends MarkupRangeOptions {
  /** Bookmark name used for reference (CT_Bookmark `@w:name`, required). */
  name: string;
  /** First column of a table-cell bookmark scope (CT_BookmarkRange `@w:colFirst`). */
  colFirst?: number;
  /** Last column of a table-cell bookmark scope (CT_BookmarkRange `@w:colLast`). */
  colLast?: number;
}

/**
 * Move revision range start (w:moveFromRangeStart / w:moveToRangeStart):
 * CT_Bookmark + author + date. `name` is optional in practice — Word does not
 * always emit one for auto-generated move ranges.
 */
export interface MoveRangeStartOptions {
  /** Move range identifier (CT_Markup `@w:id`, required). */
  id: number;
  /** Move range name (CT_Bookmark `@w:name` — required by XSD, often absent in Word). */
  name?: string;
  /** Author of the move (CT_MoveBookmark `@w:author`). */
  author?: string;
  /** Date of the move (CT_MoveBookmark `@w:date`). */
  date?: DateTime;
  /** Displacement relative to a sibling customXml (CT_MarkupRange `@w:displacedByCustomXml`). */
  displacedByCustomXml?: DisplacedByCustomXml;
  /** First column of a table-cell move range scope (CT_BookmarkRange `@w:colFirst`). */
  colFirst?: number;
  /** Last column of a table-cell move range scope (CT_BookmarkRange `@w:colLast`). */
  colLast?: number;
}

/**
 * Inline bookmark sugar: the library allocates the id and emits the paired
 * bookmarkStart/bookmarkEnd markers around `wrap` (the anchored inline content).
 * Pure markup — adds no part or relationship.
 */
export interface BookmarkOptions {
  /** Bookmark name used for reference (CT_Bookmark `@w:name`, required). */
  name: string;
  /** Anchored document content the bookmark range wraps (inline runs/text). */
  wrap?: (string | RunOptions)[];
  /** Displacement relative to a sibling customXml (CT_MarkupRange `@w:displacedByCustomXml`). */
  displacedByCustomXml?: DisplacedByCustomXml;
  /** First column of a table-cell bookmark scope (CT_BookmarkRange `@w:colFirst`). */
  colFirst?: number;
  /** Last column of a table-cell bookmark scope (CT_BookmarkRange `@w:colLast`). */
  colLast?: number;
}

/**
 * Move revision sugar for `{ moveFrom }`/`{ moveTo }`: the library allocates
 * all ids and emits the paired range markers with the moved run between them.
 * `name` is required and MUST be shared by the source/destination pair so Word
 * links them into one move; `wrap` is the moved content; `author`/`date` apply
 * to both the range start and the move run.
 */
export interface MoveRangeOptions {
  /** Author of the move (CT_MoveBookmark + move run `@w:author`, required). */
  author: string;
  /** Date of the move (CT_MoveBookmark + move run `@w:date`, required). */
  date: DateTime;
  /** Moved content carried by the move run (inline runs/text). */
  wrap?: (string | RunOptions)[];
  /** Move range name (CT_Bookmark `@w:name`, required). Share it across the
   * matching `moveFrom`/`moveTo` pair so Word links them into one move. */
  name: string;
  /** Displacement relative to a sibling customXml (CT_MarkupRange `@w:displacedByCustomXml`). */
  displacedByCustomXml?: DisplacedByCustomXml;
  /** First column of a table-cell move range scope (CT_BookmarkRange `@w:colFirst`). */
  colFirst?: number;
  /** Last column of a table-cell move range scope (CT_BookmarkRange `@w:colLast`). */
  colLast?: number;
}
