import { XmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";

import { buildEndParagraphRunProperties } from "./end-paragraph-run";
import type { IParagraphPropertiesOptions } from "./paragraph-properties";
import { buildParagraphProperties } from "./paragraph-properties";
import { TextRun } from "./run";

export interface IParagraphOptions {
    /** Simple text content for the paragraph. Creates a single TextRun. */
    readonly text?: string;
    readonly properties?: IParagraphPropertiesOptions;
    readonly children?: readonly (TextRun | XmlComponent)[];
}

/**
 * a:p — A paragraph in a text body.
 * Lazy: stores options, builds XML object in prepForXml.
 */
export class Paragraph extends XmlComponent {
    private readonly options: IParagraphOptions;

    public constructor(options: string | IParagraphOptions = {}) {
        super("a:p");
        this.options = typeof options === "string" ? { text: options } : options;
    }

    public prepForXml(context: IContext): IXmlableObject | undefined {
        const children: IXmlableObject[] = [];

        const pPr = buildParagraphProperties(this.options.properties ?? {});
        if (pPr) children.push(pPr);

        if (this.options.text) {
            const obj = new TextRun(this.options.text).prepForXml(context);
            if (obj) children.push(obj);
        }

        if (this.options.children) {
            for (const child of this.options.children) {
                const obj = child.prepForXml(context);
                if (obj) children.push(obj);
            }
        }

        children.push(buildEndParagraphRunProperties());

        return { "a:p": children };
    }
}
