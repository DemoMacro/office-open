import { XmlComponent } from "@file/xml-components";

/**
 * a:cNvPicPr — Non-visual picture drawing properties.
 * Uses a: prefix (DrawingML type) but referenced via p:cNvPicPr in PML context.
 */
export class NonVisualPictureProperties extends XmlComponent {
    public constructor() {
        super("a:cNvPicPr");
    }
}
