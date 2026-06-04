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
import type { IXmlableObject } from "@file/xml-components";
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
  private readonly options: SmartTagRunOptions;

  public constructor(options: SmartTagRunOptions) {
    super("w:smartTag");
    this.options = options;

    const attrs: Record<string, string> = {};
    if (options.uri !== undefined) attrs["w:uri"] = options.uri;
    if (options.element !== undefined) attrs["w:element"] = options.element;
    if (Object.keys(attrs).length > 0) {
      this.root.push({ _attr: attrs });
    }

    if (options.properties && options.properties.length > 0) {
      this.root.push(new SmartTagProperties(options.properties));
    }
  }

  public override prepForXml(
    context: import("@file/xml-components").Context,
  ): IXmlableObject | undefined {
    const attrs: Record<string, string> = {};
    if (this.options.uri !== undefined) attrs["w:uri"] = this.options.uri;
    if (this.options.element !== undefined) attrs["w:element"] = this.options.element;

    const inner: (IXmlableObject | { _attr: Record<string, string> })[] = [];
    if (Object.keys(attrs).length > 0) inner.push({ _attr: attrs });

    if (this.options.properties && this.options.properties.length > 0) {
      const attrChildren: IXmlableObject[] = [];
      for (const a of this.options.properties) {
        attrChildren.push({
          "w:attr": { _attr: { "w:uri": a.uri, "w:name": a.name, "w:val": a.val } },
        });
      }
      inner.push({ "w:smartTagPr": attrChildren });
    }

    if (this.options.children) {
      for (const child of this.options.children) {
        if (child instanceof BaseXmlComp) {
          const obj = child.prepForXml(context);
          if (obj) inner.push(obj);
        }
      }
    }

    return inner.length > 0 ? { "w:smartTag": inner } : { "w:smartTag": {} };
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
