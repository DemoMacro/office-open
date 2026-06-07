/**
 * Context objects for the descriptor read/write pipeline.
 *
 * @module
 */

import type { Element as XmlElement } from "@office-open/xml";

/** Context passed during stringify (write path). */
export interface WriteContext {
  /** Register a relationship and return its rId. */
  addRelationship(type: string, target: string, mode?: string): string;
  /** Add a media file and return its reference. */
  addMedia(data: Uint8Array, type: string): string;
}

/** Context passed during parse (parse path). */
export interface ReadContext {
  /** Resolve a relationship rId to its target path. */
  resolveRelationship(rId: string): string | undefined;
  /** Get a parsed XML part by path. */
  getPart(path: string): XmlElement | undefined;
  /** Get raw binary data (images, media, etc.) by path. */
  getRaw(path: string): Uint8Array | undefined;
}
