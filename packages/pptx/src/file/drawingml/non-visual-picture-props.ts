import { EmptyElement } from "@file/xml-components";

/**
 * p:cNvPicPr — Non-visual picture drawing properties (within p:pic context).
 */
export class NonVisualPictureProperties extends EmptyElement {
  public constructor() {
    super("p:cNvPicPr");
  }
}
