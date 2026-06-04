/**
 * SmartTagRun module for WordprocessingML documents.
 *
 * Smart tags identify recognized text (names, dates, addresses, etc.)
 * and associate it with actions.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_SmartTagRun
 *
 * @module
 */
import {
  BaseXmlComponent as BaseXmlComp,
  BuilderElement,
  XmlComponent,
} from "@file/xml-components";

import type { ParagraphChild } from "../paragraph";

/**
 * Options for creating a SmartTagRun.
 */
export interface SmartTagRunOptions {
  /** Namespace URI of the smart tag */
  readonly uri?: string;
  /** Element name within the namespace */
  readonly element: string;
  /** Attributes for w:smartTagPr (namespace:uri, name:val pairs) */
  readonly properties?: readonly {
    readonly uri: string;
    readonly name: string;
    readonly val: string;
  }[];
  /** Inline content children */
  readonly children?: readonly ParagraphChild[];
}

/**
 * Represents a smart tag run (w:smartTag) in a WordprocessingML document.
 *
 * Smart tags wrap inline content to provide contextual actions.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_SmartTagRun
 *
 * @example
 * ```typescript
 * new SmartTagRun({
 *   uri: "urn:schemas-microsoft-com:office:smarttags",
 *   element: "date",
 *   children: [new TextRun("2024-01-01")],
 * });
 * ```
 */
export class SmartTagRun extends XmlComponent {
  public constructor(options: SmartTagRunOptions) {
    super("w:smartTag");

    const attrs: Record<string, string> = {};
    if (options.uri !== undefined) attrs["w:uri"] = options.uri;
    if (options.element !== undefined) attrs["w:element"] = options.element;
    if (Object.keys(attrs).length > 0) {
      this.root.push({ _attr: attrs });
    }

    if (options.properties && options.properties.length > 0) {
      this.root.push(new SmartTagProperties(options.properties));
    }

    if (options.children) {
      for (const child of options.children) {
        if (child instanceof BaseXmlComp) {
          this.root.push(child);
        }
      }
    }
  }
}

/**
 * Smart tag properties (w:smartTagPr).
 *
 * Contains w:attr children mapping namespace/name/value triples.
 */
class SmartTagProperties extends XmlComponent {
  public constructor(
    attrs?: readonly { readonly uri: string; readonly name: string; readonly val: string }[],
  ) {
    super("w:smartTagPr");

    if (attrs) {
      for (const a of attrs) {
        this.root.push(
          new BuilderElement({
            name: "w:attr",
            attributes: {
              uri: { key: "w:uri", value: a.uri },
              name: { key: "w:name", value: a.name },
              val: { key: "w:val", value: a.val },
            },
          }),
        );
      }
    }
  }
}
