import type { Paragraph } from "@file/paragraph";
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

export const createTextBoxContent = (children: readonly Paragraph[]): XmlComponent =>
    new BuilderElement({
        children: [...children],
        name: "wps:txbxContent",
    });
