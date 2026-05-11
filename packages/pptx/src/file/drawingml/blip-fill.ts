import { XmlComponent } from "@file/xml-components";
import { createBlip, Stretch } from "@office-open/core/drawingml";

/**
 * a:blipFill — Image fill with stretch mode.
 * Uses core createBlip for the blip element.
 */
export class BlipFill extends XmlComponent {
    public constructor(fileName: string) {
        super("a:blipFill");
        this.root.push(createBlip({ referenceId: fileName }));
        this.root.push(new Stretch());
    }
}
