import { XmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";

import type { IRunPropertiesOptions } from "./run-properties";
import { RunProperties } from "./run-properties";

export interface IRunOptions extends IRunPropertiesOptions {
    readonly text?: string;
}

/**
 * a:r — A run of text with properties.
 * Lazy: stores options, builds XML object in prepForXml.
 */
export class Run extends XmlComponent {
    private readonly options: IRunOptions;

    public constructor(options: IRunOptions = {}) {
        super("a:r");
        this.options = options;
    }

    public prepForXml(context: IContext): IXmlableObject | undefined {
        const children: IXmlableObject[] = [];

        if (RunProperties.hasProperties(this.options)) {
            const rPr = new RunProperties(this.options);
            const rPrObj = rPr.prepForXml(context);
            if (rPrObj) children.push(rPrObj);
        }

        if (this.options.text) {
            children.push({ "a:t": [this.options.text] });
        }

        return {
            "a:r":
                children.length === 0
                    ? {}
                    : children.length === 1 && "_attr" in children[0]
                      ? children[0]
                      : children,
        };
    }
}
