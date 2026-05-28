/**
 * Document wrapper module for WordprocessingML documents.
 *
 * This module provides wrappers that combine the main document/header/footer
 * views with their associated relationships, enabling proper management of
 * document parts and their references.
 *
 * @module
 */
import { Document } from "./document";
import type { DocumentOptions } from "./document";
import type { Endnotes } from "./endnotes";
import type { Footer } from "./footer/footer";
import type { FootNotes } from "./footnotes";
import type { Header } from "./header/header";
import { Relationships } from "./relationships";
import type { XmlComponent } from "./xml-components";

/**
 * Interface for document view wrappers.
 *
 * ViewWrappers combine a document part (view) with its relationships,
 * providing a unified interface for managing document components.
 *
 * @property view - The document part (Document, Header, Footer, etc.)
 * @property relationships - The relationships associated with this view
 */
export interface ViewWrapper {
  readonly view: Document | Footer | Header | FootNotes | Endnotes | XmlComponent;
  readonly relationships: Relationships;
}

/**
 * Wrapper for the main document body.
 *
 * DocumentWrapper combines the main Document view with its Relationships,
 * managing the primary content of the .docx file along with references to
 * images, hyperlinks, and other linked resources.
 *
 * @example
 * ```typescript
 * const wrapper = new DocumentWrapper({
 *   sections: [{
 *     children: [new Paragraph("Hello World")],
 *   }],
 * });
 * ```
 */
export class DocumentWrapper implements ViewWrapper {
  private readonly document: Document;
  public readonly relationships: Relationships;

  public constructor(options: DocumentOptions) {
    this.document = new Document(options);
    this.relationships = new Relationships();
  }

  public get view(): Document {
    return this.document;
  }
}
