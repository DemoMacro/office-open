import { EmptyElement } from "@file/xml-components";

/**
 * p:cNvSpPr — Non-visual shape properties.
 * Uses p: prefix in PresentationML context.
 */
export class NonVisualShapeProperties extends EmptyElement {
  public constructor() {
    super("p:cNvSpPr");
  }
}
