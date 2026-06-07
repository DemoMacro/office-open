/**
 * Parse context for PPTX documents.
 *
 * @module
 */
import type { PptxDocument } from "../parse";

export class ParseContext {
  constructor(
    public pptx: PptxDocument,
    /** Slide relationship ID → path, parsed from slide's _rels file */
    public slideRels: Map<string, string>,
  ) {}
}
