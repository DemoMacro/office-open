/**
 * Bookmark types for WordprocessingML documents.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd — CT_Markup, CT_MarkupRange,
 * CT_BookmarkRange, CT_Bookmark.
 *
 * @module
 */

/**
 * ST_DisplacedByCustomXml — whether a range marker is displaced before or
 * after a sibling customXml element. Shared by every CT_MarkupRange derivative
 * (bookmark, move range, comment range, …).
 */
export type DisplacedByCustomXml = "before" | "after";

/**
 * Options for a bookmark end (w:bookmarkEnd).
 *
 * Maps to CT_MarkupRange (CT_Markup `@id` + `@displacedByCustomXml`). Unlike
 * bookmarkStart, the end marker carries neither a name nor a column range.
 *
 * Reference: wml.xsd CT_MarkupRange, EG_RangeMarkupElements
 */
export interface BookmarkEndOptions {
  /** Bookmark identifier (CT_Markup @w:id, required). */
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
export interface BookmarkStartOptions extends BookmarkEndOptions {
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
