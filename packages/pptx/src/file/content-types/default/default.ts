import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

interface IDefaultAttributes {
    readonly contentType: string;
    readonly extension?: string;
}

export const createDefault = (contentType: string, extension?: string): XmlComponent =>
    new BuilderElement<IDefaultAttributes>({
        attributes: {
            contentType: { key: "ContentType", value: contentType },
            extension: { key: "Extension", value: extension },
        },
        name: "Default",
    });
