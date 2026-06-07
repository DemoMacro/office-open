/**
 * Parse context module.
 *
 * Provides the shared context object used throughout the DOCX parsing pipeline.
 * Holds references to the parsed DocxDocument and cached style/numbering data.
 *
 * @module
 */
import type { Element } from "@office-open/xml";

import type { DocxDocument } from "../parse";

export class ParseContext {
  constructor(
    public docx: DocxDocument,
    public styleCache: Map<string, Element>,
    public numberingCache: Map<string, Element>,
  ) {}
}
