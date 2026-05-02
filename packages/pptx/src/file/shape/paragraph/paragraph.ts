import { XmlComponent } from "@file/xml-components";

import { EndParagraphRunProperties } from "./end-paragraph-run";
import type { IParagraphPropertiesOptions } from "./paragraph-properties";
import { ParagraphProperties } from "./paragraph-properties";
import { Run } from "./run";

export interface IParagraphOptions {
    readonly properties?: IParagraphPropertiesOptions;
    readonly children?: readonly (Run | XmlComponent)[];
}

/**
 * a:p — A paragraph in a text body.
 */
export class Paragraph extends XmlComponent {
    public constructor(options: IParagraphOptions = {}) {
        super("a:p");
        this.root.push(new ParagraphProperties(options.properties ?? {}));
        if (options.children) {
            this.root.push(...options.children);
        }
        this.root.push(new EndParagraphRunProperties());
    }
}
