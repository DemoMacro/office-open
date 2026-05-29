import { XmlComponent } from "@file/xml-components";
import { createBlip, Stretch } from "@office-open/core/drawingml";

/**
 * p:blipFill — Image fill with stretch mode (within p:pic context).
 * Uses core createBlip for the blip element.
 */
export class BlipFill extends XmlComponent {
  public constructor(fileName: string) {
    super("p:blipFill");
    this.root.push(createBlip({ referenceId: fileName }));
    this.root.push(new Stretch());
  }
}
