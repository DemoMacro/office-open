/**
 * Footnotes wrapper module for WordprocessingML documents.
 *
 * This module provides a wrapper for managing footnotes and their
 * associated relationships in a document.
 *
 * Reference: http://officeopenxml.com/WPfootnotes.php
 *
 * @module
 */
import type { ViewWrapper } from "./document-wrapper";
import { FootNotes } from "./footnotes/footnotes";
import { Relationships } from "./relationships";

/**
 * Wrapper class for managing footnotes in a document.
 *
 * Encapsulates the footnotes collection and its relationships,
 * implementing the ViewWrapper interface for consistent access.
 *
 * @example
 * ```typescript
 * const wrapper = new FootnotesWrapper();
 * const footnotes = wrapper.view;
 * const relationships = wrapper.relationships;
 * ```
 */
export class FootnotesWrapper implements ViewWrapper {
  private footnotes: FootNotes;
  public relationships: Relationships;

  public constructor() {
    this.footnotes = new FootNotes();
    this.relationships = new Relationships();
  }

  public get view(): FootNotes {
    return this.footnotes;
  }
}
