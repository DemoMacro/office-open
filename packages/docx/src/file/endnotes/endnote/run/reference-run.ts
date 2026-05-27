import { Run } from "@file/paragraph/run";
import { XmlComponent } from "@file/xml-components";

export class EndnoteIdReference extends XmlComponent {
  public constructor(id: number) {
    super("w:endnoteReference");

    this.root.push({ _attr: { "w:id": id } });
  }
}

export class EndnoteReferenceRun extends Run {
  public constructor(id: number) {
    super({ style: "EndnoteReference" });

    this.extraChildren.push(new EndnoteIdReference(id));
  }
}
