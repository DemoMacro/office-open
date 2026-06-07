/**
 * Header and footer module for WordprocessingML documents.
 *
 * This module provides classes for creating headers and footers
 * that can be attached to document sections.
 *
 * Reference: http://officeopenxml.com/WPSectionHeaderFooters.php
 *
 * @module
 */
import type { SectionChild } from "./section-child";

/**
 * Options for creating a header or footer.
 *
 * @see {@link Header}
 * @see {@link Footer}
 */
export interface HeaderOptions {
  /** The content elements (paragraphs, tables, etc.) for the header/footer */
  children: SectionChild[];
}

/**
 * Represents a document header.
 *
 * Headers appear at the top of each page in a section and can contain
 * paragraphs, tables, images, and other content.
 *
 * @publicApi
 *
 * @example
 * ```typescript
 * const header = new Header({
 *   children: [
 *     new Paragraph({ children: [new TextRun("Company Name")] }),
 *   ],
 * });
 * ```
 */
export class Header {
  public options: HeaderOptions;

  public constructor(options: HeaderOptions = { children: [] }) {
    this.options = options;
  }
}

/**
 * Represents a document footer.
 *
 * Footers appear at the bottom of each page in a section and can contain
 * paragraphs, tables, images, and other content.
 *
 * @publicApi
 *
 * @example
 * ```typescript
 * const footer = new Footer({
 *   children: [
 *     new Paragraph({ children: [new TextRun("Page "), PageNumber.CURRENT] }),
 *   ],
 * });
 * ```
 */
export class Footer {
  public options: HeaderOptions;

  public constructor(options: HeaderOptions = { children: [] }) {
    this.options = options;
  }
}
