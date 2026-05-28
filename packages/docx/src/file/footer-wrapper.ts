/**
 * Footer wrapper module for WordprocessingML documents.
 *
 * This module provides a wrapper for document footers, managing footer content
 * along with their relationships and media. Footers can be configured for
 * different pages (first, even, odd, default).
 *
 * Reference: http://officeopenxml.com/WPfooter.php
 *
 * @module
 */
import type { XmlComponent } from "@file/xml-components";

import type { HeaderFooterReferenceType } from "./document";
import type { ViewWrapper } from "./document-wrapper";
import type { FileChild } from "./file-child";
import { Footer } from "./footer/footer";
import type { Media } from "./media";
import { Relationships } from "./relationships";

/**
 * Configuration for a document footer.
 *
 * @property footer - The FooterWrapper instance containing the footer content
 * @property type - The footer type (default, first page, even pages)
 */
export interface DocumentFooter {
  readonly footer: FooterWrapper;
  readonly type: (typeof HeaderFooterReferenceType)[keyof typeof HeaderFooterReferenceType];
}

/**
 * Wrapper for document footers.
 *
 * FooterWrapper combines a Footer view with its Relationships and Media,
 * enabling footers to contain paragraphs, tables, images, and hyperlinks.
 * Each section can have multiple footers for different page types.
 *
 * Reference: http://officeopenxml.com/WPfooter.php
 *
 * @example
 * ```typescript
 * const footerWrapper = new FooterWrapper(media, 1);
 * footerWrapper.add(new Paragraph("Page Footer"));
 * footerWrapper.add(new Table({
 *   rows: [new TableRow({ children: [new TableCell({ children: [new Paragraph("Cell")] })] })],
 * }));
 * ```
 */
export class FooterWrapper implements ViewWrapper {
  private readonly footer: Footer;
  public readonly relationships: Relationships;
  public readonly media: Media;

  public constructor(media: Media, referenceId: number, initContent?: XmlComponent) {
    this.media = media;
    this.footer = new Footer(referenceId, initContent);
    this.relationships = new Relationships();
  }

  public add(item: FileChild): void {
    this.footer.add(item);
  }

  public addChildElement(childElement: XmlComponent): void {
    this.footer.addChildElement(childElement);
  }

  public get view(): Footer {
    return this.footer;
  }
}
