import { XmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { xml } from "@office-open/xml";

import type { ParagraphPropertiesOptions } from "./paragraph-properties";
import { buildParagraphProperties } from "./paragraph-properties";
import { TextRun } from "./run";
import type { RunOptions } from "./run";

export interface ParagraphOptions {
  /** Simple text content for the paragraph. Creates a single TextRun. */
  readonly text?: string;
  readonly properties?: ParagraphPropertiesOptions;
  readonly children?: readonly (TextRun | RunOptions | XmlComponent)[];
}

/**
 * a:p — A paragraph in a text body.
 * Lazy: stores options, builds XML in toXml().
 */
export class Paragraph extends XmlComponent {
  private readonly options: ParagraphOptions;

  public constructor(options: string | ParagraphOptions = {}) {
    super("a:p");
    this.options = typeof options === "string" ? { text: options } : options;
  }

  public override toXml(context: Context): string {
    const parts: string[] = [];

    const pPr = buildParagraphProperties(this.options.properties ?? {});
    if (pPr) parts.push(xml(pPr));

    if (this.options.text) {
      parts.push(new TextRun(this.options.text).toXml(context));
    }

    if (this.options.children) {
      for (const rawChild of this.options.children) {
        const child =
          rawChild instanceof TextRun || rawChild instanceof XmlComponent
            ? rawChild
            : new TextRun(rawChild);
        parts.push(child.toXml(context));
      }
    }

    parts.push('<a:endParaRPr lang="en-US"/>');

    const body = parts.join("");
    return body ? `<a:p>${body}</a:p>` : "<a:p/>";
  }
}
