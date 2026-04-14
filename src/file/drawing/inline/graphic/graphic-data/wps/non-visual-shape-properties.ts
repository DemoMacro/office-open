import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

export interface INonVisualShapePropertiesOptions {
    readonly txBox: string;
}

export const createNonVisualShapeProperties = (
    options: INonVisualShapePropertiesOptions = { txBox: "1" },
): XmlComponent =>
    new BuilderElement<{ readonly txBox: string }>({
        attributes: {
            txBox: { key: "txBox", value: options.txBox },
        },
        name: "wps:cNvSpPr",
    });
