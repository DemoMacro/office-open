import type { Paragraph } from "@file/paragraph";
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

import { createTextBoxContent } from "./text-box-content";

export const createWpsTextBox = (children: readonly Paragraph[]): XmlComponent =>
    new BuilderElement({
        children: [createTextBoxContent(children)],
        name: "wps:txbx",
    });
