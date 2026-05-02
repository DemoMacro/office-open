import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

interface IOverrideAttributes {
    readonly contentType: string;
    readonly partName?: string;
}

export const createOverride = (contentType: string, partName?: string): XmlComponent =>
    new BuilderElement<IOverrideAttributes>({
        attributes: {
            contentType: { key: "ContentType", value: contentType },
            partName: { key: "PartName", value: partName },
        },
        name: "Override",
    });
