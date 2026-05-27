/**
 * Bidirectional text override elements for WordprocessingML documents.
 *
 * These elements allow overriding the default text direction for inline content,
 * used for embedding LTR text inside RTL paragraphs or vice versa.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_DirContentRun, CT_BdoContentRun
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

import type { ParagraphChild } from "../paragraph";

/**
 * Options for directional text override.
 */
export interface DirOptions {
  /** Array of paragraph children inside the direction override */
  readonly children: readonly ParagraphChild[];
  /** Text direction: "ltr" or "rtl" */
  readonly val: "ltr" | "rtl";
}

/**
 * Overrides the default bidirectional text direction for inline content.
 *
 * Unlike `w:bdo`, this respects the Unicode bidirectional algorithm
 * and only provides a hint.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, w:dir (CT_DirContentRun)
 *
 * @example
 * ```typescript
 * new Dir({ val: "rtl", children: [new TextRun("مرحبا")] });
 * ```
 */
export class Dir extends XmlComponent {
  public constructor(options: DirOptions) {
    super("w:dir");
    this.root.push({ _attr: { "w:val": options.val } });
    for (const child of options.children) {
      this.root.push(child);
    }
  }
}

/**
 * Strong bidirectional override for inline content.
 *
 * Forces the specified direction regardless of the Unicode bidi algorithm.
 * Contains full paragraph content (runs, breaks, etc.).
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, w:bdo (CT_BdoContentRun)
 *
 * @example
 * ```typescript
 * new Bdo({ val: "ltr", children: [new TextRun("Hello")] });
 * ```
 */
export class Bdo extends XmlComponent {
  public constructor(options: DirOptions) {
    super("w:bdo");
    this.root.push({ _attr: { "w:val": options.val } });
    for (const child of options.children) {
      this.root.push(child);
    }
  }
}
