import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

export interface NonVisualShapePropertiesOptions {
  txBox: string;
}

export const createNonVisualShapeProperties = (
  options: NonVisualShapePropertiesOptions = { txBox: "1" },
): XmlComponent =>
  new BuilderElement<{ txBox: string }>({
    attributes: {
      txBox: { key: "txBox", value: options.txBox },
    },
    name: "wps:cNvSpPr",
  });
