import type { FileChild } from "@file/file-child";
import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

export const createTextBoxContent = (children: FileChild[]): XmlComponent =>
  new BuilderElement({
    children: children,
    name: "wps:txbxContent",
  });
