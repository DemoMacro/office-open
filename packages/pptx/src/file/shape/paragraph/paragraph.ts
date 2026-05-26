import { XmlComponent } from "@file/xml-components";
import type { Context, IXmlableObject } from "@file/xml-components";

import { buildEndParagraphRunProperties } from "./end-paragraph-run";
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
 * Lazy: stores options, builds XML object in prepForXml.
 */
export class Paragraph extends XmlComponent {
  private readonly options: ParagraphOptions;

  public constructor(options: string | ParagraphOptions = {}) {
    super("a:p");
    this.options = typeof options === "string" ? { text: options } : options;
  }

  public prepForXml(context: Context): IXmlableObject | undefined {
    const children: IXmlableObject[] = [];

    const pPr = buildParagraphProperties(this.options.properties ?? {});
    if (pPr) children.push(pPr);

    if (this.options.text) {
      const obj = new TextRun(this.options.text).prepForXml(context);
      if (obj) children.push(obj);
    }

    if (this.options.children) {
      for (const rawChild of this.options.children) {
        const child =
          rawChild instanceof TextRun || rawChild instanceof XmlComponent
            ? rawChild
            : new TextRun(rawChild);
        const obj = child.prepForXml(context);
        if (obj) children.push(obj);
      }
    }

    children.push(buildEndParagraphRunProperties());

    return { "a:p": children };
  }
}
