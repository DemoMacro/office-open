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
 * @property id - Unique identifier for this change (must be unique within the document)
 * @property author - Name of the author who made the change
 * @property date - Date and time when the change was made (ISO 8601 format)
 */
export interface ChangedAttributesProperties {
  /** Unique identifier for this change (must be unique within the document) */
  readonly id: number;
  /** Name of the author who made the change */
  readonly author: string;
  /** Date and time when the change was made (ISO 8601 format) */
  readonly date: string;
}
