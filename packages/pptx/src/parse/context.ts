/**
 * Parse context for PPTX documents.
 *
 * @module
 */
import type { PptxDocument } from "../parse";

export class ParseContext {
  constructor(
    readonly pptx: PptxDocument,
    /** Slide relationship ID → path, parsed from slide's _rels file */
    readonly slideRels: Map<string, string>,
  ) {}
}
