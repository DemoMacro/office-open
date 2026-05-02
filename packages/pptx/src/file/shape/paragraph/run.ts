import { XmlComponent } from "@file/xml-components";

import type { IRunPropertiesOptions } from "./run-properties";
import { RunProperties } from "./run-properties";
import { Text } from "./text";

export interface IRunOptions extends IRunPropertiesOptions {
    readonly text?: string;
}

/**
 * a:r — A run of text with properties.
 */
export class Run extends XmlComponent {
    public constructor(options: IRunOptions = {}) {
        super("a:r");
        if (RunProperties.hasProperties(options)) {
            this.root.push(new RunProperties(options));
        }
        if (options.text) {
            this.root.push(new Text(options.text));
        }
    }
}
