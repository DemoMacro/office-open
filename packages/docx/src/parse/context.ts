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
    readonly docx: DocxDocument,
    readonly styleCache: Map<string, Element>,
    readonly numberingCache: Map<string, Element>,
  ) {}
}
