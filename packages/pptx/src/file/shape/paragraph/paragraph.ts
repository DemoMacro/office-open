import { XmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";

import { buildEndParagraphRunProperties } from "./end-paragraph-run";
import type { IParagraphPropertiesOptions } from "./paragraph-properties";
import { buildParagraphProperties } from "./paragraph-properties";
import { Run } from "./run";

export interface IParagraphOptions {
    readonly properties?: IParagraphPropertiesOptions;
    readonly children?: readonly (Run | XmlComponent)[];
}

/**
 * a:p — A paragraph in a text body.
 * Lazy: stores options, builds XML object in prepForXml.
 */
export class Paragraph extends XmlComponent {
    private readonly options: IParagraphOptions;

    public constructor(options: IParagraphOptions = {}) {
        super("a:p");
        this.options = options;
    }

    public prepForXml(context: IContext): IXmlableObject | undefined {
        const children: IXmlableObject[] = [];

        const pPr = buildParagraphProperties(this.options.properties ?? {});
        if (pPr) children.push(pPr);

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
