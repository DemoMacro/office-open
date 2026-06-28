/**
 * Bookmark types for WordprocessingML documents.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd — CT_Markup, CT_MarkupRange,
 * CT_BookmarkRange, CT_Bookmark.
 *
 * @module
 */

import type { RunOptions } from "@parts/paragraph/run/run";

/**
 * ST_DisplacedByCustomXml — whether a range marker is displaced before or
 * after a sibling customXml element. Shared by every CT_MarkupRange derivative
 * (bookmark, move range, comment range, …).
 */
export type DisplacedByCustomXml = "before" | "after";

/**
 * Options for a CT_MarkupRange end marker — `@id` + `@displacedByCustomXml`.
 *
 * Shared by every range end marker that derives from CT_MarkupRange:
 * w:bookmarkEnd, w:commentRangeStart/End, w:moveFromRangeEnd, w:moveToRangeEnd.
 * None of these carry a name or column range (only the start markers do).
 *
 * Reference: wml.xsd CT_MarkupRange, EG_RangeMarkupElements
 */
export interface MarkupRangeOptions {
  /** Marker identifier (CT_Markup @w:id, required). */
  id: number;
  /** Displacement relative to a sibling customXml (CT_MarkupRange @w:displacedByCustomXml). */
  displacedByCustomXml?: DisplacedByCustomXml;
}

/**
 * Options for a bookmark start (w:bookmarkStart).
 *
 * Maps to CT_Bookmark = CT_BookmarkRange (colFirst/colLast) + name. The column
 * range scopes a bookmark to specific table columns so Word preserves the exact
 * cell span on round-trip rather than recomputing it.
 *
 * Reference: wml.xsd CT_Bookmark, CT_BookmarkRange
 */
export interface BookmarkStartOptions extends MarkupRangeOptions {
  /** Bookmark name used for reference (CT_Bookmark @w:name, required). */
  name: string;
  /** First column of a table-cell bookmark scope (CT_BookmarkRange @w:colFirst). */
  colFirst?: number;
  /** Last column of a table-cell bookmark scope (CT_BookmarkRange @w:colLast). */
  colLast?: number;
}

/**
 * Options for a move revision range start (w:moveFromRangeStart / w:moveToRangeStart).
 *
 * Maps to CT_MoveBookmark = CT_Bookmark (name, colFirst, colLast,
 * displacedByCustomXml) + author + date. `name` is optional in practice:
 * Word does not always emit one for auto-generated move ranges.
 *
 * Reference: wml.xsd CT_MoveBookmark, EG_RangeMarkupElements
 */
export interface MoveRangeStartOptions {
  /** Move range identifier (CT_Markup @w:id, required). */
  id: number;
  /** Move range name (CT_Bookmark @w:name — required by XSD, often absent in Word). */
  name?: string;
  /** Author of the move (CT_MoveBookmark @w:author). */
  author?: string;
  /** Date of the move (CT_MoveBookmark @w:date). */
  date?: string;
  /** Displacement relative to a sibling customXml (CT_MarkupRange @w:displacedByCustomXml). */
  displacedByCustomXml?: DisplacedByCustomXml;
  /** First column of a table-cell move range scope (CT_BookmarkRange @w:colFirst). */
  colFirst?: number;
  /** Last column of a table-cell move range scope (CT_BookmarkRange @w:colLast). */
  colLast?: number;
}

/**
 * A bookmark authored as a single inline paragraph child. The library allocates
 * the bookmark id and emits the paired bookmarkStart/bookmarkEnd with one shared
 * id — the caller never touches an id or pairs markers.
 *
 * `wrap` is the anchored document content the bookmark range wraps (inline
 * runs/text, emitted between the markers). Bookmarks are pure markup: they add
 * no part or relationship, only the two markers in document.xml.
 *
 * Reference: wml.xsd CT_Bookmark, CT_BookmarkRange, EG_RangeMarkupElements.
 */
export interface BookmarkOptions {
  /** Bookmark name used for reference (CT_Bookmark @w:name, required). */
  name: string;
  /** Anchored document content the bookmark range wraps (inline runs/text). */
  wrap?: (string | RunOptions)[];
  /** Displacement relative to a sibling customXml (CT_MarkupRange @w:displacedByCustomXml). */
  displacedByCustomXml?: DisplacedByCustomXml;
  /** First column of a table-cell bookmark scope (CT_BookmarkRange @w:colFirst). */
  colFirst?: number;
  /** Last column of a table-cell bookmark scope (CT_BookmarkRange @w:colLast). */
  colLast?: number;
}

/**
 * A move revision authored as a single inline paragraph child. The library
 * allocates the range id and the move-run id, then emits the paired range
 * markers with the moved run between them — the caller never touches an id.
 *
 * Used by both the `{ moveFrom }` (source) and `{ moveTo }` (destination)
 * sugars. `wrap` is the moved content carried by the move run. `author` and
 * `date` apply to both the range start (CT_MoveBookmark) and the move run
 * (CT_TrackChange). `name` is required by XSD (CT_Bookmark); the source
 * (`moveFrom`) and destination (`moveTo`) of one logical move MUST share a name
 * so Word pairs them as a single move.
 *
 * Reference: wml.xsd CT_MoveBookmark, CT_TrackChange, EG_RangeMarkupElements.
 */
export interface MoveRangeOptions {
  /** Author of the move (CT_MoveBookmark + move run @w:author, required). */
  author: string;
  /** Date of the move (CT_MoveBookmark + move run @w:date, required). */
  date: string;
  /** Moved content carried by the move run (inline runs/text). */
  wrap?: (string | RunOptions)[];
  /** Move range name (CT_Bookmark @w:name, required). Share it across the
   * matching `moveFrom`/`moveTo` pair so Word links them into one move. */
  name: string;
  /** Displacement relative to a sibling customXml (CT_MarkupRange @w:displacedByCustomXml). */
  displacedByCustomXml?: DisplacedByCustomXml;
  /** First column of a table-cell move range scope (CT_BookmarkRange @w:colFirst). */
  colFirst?: number;
  /** Last column of a table-cell move range scope (CT_BookmarkRange @w:colLast). */
  colLast?: number;
}
