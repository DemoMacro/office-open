import { EndnoteReference } from "@file/paragraph/run/empty-children";
import { Run } from "@file/paragraph/run/run";

export class EndnoteRefRun extends Run {
  public constructor() {
    super({
      style: "EndnoteReference",
    });

    this.extraChildren.push(new EndnoteReference());
  }
}
