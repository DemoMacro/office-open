/**
 * Context objects for the descriptor read/write pipeline.
 *
 * @module
 */

import type { Element as XmlElement } from "@office-open/xml";

/** Target for a DrawingML text hyperlink (external URL or internal slide jump). */
export interface HyperlinkTarget {
  /** External URL (mutually exclusive with slide). */
  url?: string;
  /**
   * Internal slide number, 1-based (mutually exclusive with url). PPTX only —
   * formats without slides reject it at generate time.
   */
  slide?: number;
  /** Optional tooltip. */
  tooltip?: string;
}

/** Context passed during stringify (write path). */
export interface WriteContext {
  /** Register a relationship and return its rId. */
  addRelationship(type: string, target: string, mode?: string): string;
  /** Add a media file and return its reference. */
  addMedia(data: Uint8Array, type: string): string;
  /**
   * Register a DrawingML text hyperlink (a:hlinkClick on runs). Formats that
   * don't emit DrawingML text hyperlinks (DOCX uses w:hyperlink) implement
   * this as a no-op.
   */
  addHyperlink(key: string, target: HyperlinkTarget): void;
}

/** Context passed during parse (parse path). */
export interface ReadContext {
  /** Resolve a relationship rId to its target path. */
  resolveRelationship(rId: string): string | undefined;
  /**
   * Relationship kind of an OLE embedding rId in the current part
   * ("oleObject" | "package"). Optional — only document formats that
   * distinguish the two embedding styles implement it.
   */
  resolveEmbeddingType?(rId: string): "oleObject" | "package" | undefined;
  /**
   * External image source URL of an a:blip @r:link rId in the current part.
   * Optional — formats without linked-image support implement nothing.
   */
  resolveExternalImage?(rId: string): string | undefined;
  /** Get a parsed XML part by path. */
  getPart(path: string): XmlElement | undefined;
  /** Get raw binary data (images, media, etc.) by path. */
  getRaw(path: string): Uint8Array | undefined;
}
