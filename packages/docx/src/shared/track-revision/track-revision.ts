import type { DateTime } from "@office-open/core";
/**
 * Track Revision module for WordprocessingML documents.
 *
 * This module provides support for tracking document revisions
 * (insertions, deletions, and formatting changes).
 *
 * Reference: http://officeopenxml.com/WPtrackChanges.php
 *
 * @module
 */

/**
 * Properties for a tracked change element.
 *
 * These properties identify the change and its author for revision tracking.
 *
 * @property id - Revision marker id (document-unique; a fresh document may omit
 *   it and get a library-assigned value — round-tripped sources keep theirs)
 * @property author - Name of the author who made the change
 * @property date - Date and time when the change was made (ISO 8601 format)
 */
export interface ChangedProperties {
  /**
   * Revision marker id, unique within the document. No other element or
   * attribute references it, so fresh input may omit it — the library
   * assigns one (`autoRevisionId`); explicit duplicates still get
   * renumbered by the body emit pass.
   */
  id?: number;
  /** Name of the author who made the change */
  author: string;
  /** Date and time when the change was made (ISO 8601 format) */
  date: DateTime;
}

// Revision ids only need to be document-unique (nothing references them), so a
// module-level counter suffices — monotonic across generates, never colliding
// within one document. Explicit user ids that collide are renumbered by the
// dedupe pass in body.ts.
let revisionIdCounter = 0;

/** Allocate the next revision marker id for a fresh document. */
export function autoRevisionId(): number {
  return ++revisionIdCounter;
}
