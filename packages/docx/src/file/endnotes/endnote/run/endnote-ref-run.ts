import { EndnoteReference, Run } from "@file/paragraph";

export class EndnoteRefRun extends Run {
    public constructor() {
        super({
            style: "EndnoteReference",
        });

        this.extraChildren.push(new EndnoteReference());
    }
}
