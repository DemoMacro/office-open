import { XmlComponent } from "@file/xml-components";

/**
 * a:t — Text content within a run.
 */
export class Text extends XmlComponent {
  public constructor(value: string) {
    super("a:t");
    this.root.push(value);
  }
}
