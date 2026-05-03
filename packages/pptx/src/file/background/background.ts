import type { FillOptions } from "@file/drawingml/fill";
import { buildFill } from "@file/drawingml/fill";
import { XmlComponent } from "@file/xml-components";

export interface IBackgroundOptions {
    readonly fill?: FillOptions;
}

/**
 * p:bg — Slide background.
 */
export class Background extends XmlComponent {
    public constructor(options: IBackgroundOptions = {}) {
        super("p:bg");
        this.root.push(new BackgroundProperties(options.fill));
    }
}

class BackgroundProperties extends XmlComponent {
    public constructor(fill?: FillOptions) {
        super("p:bgPr");
        this.root.push(fill !== undefined ? buildFill(fill) : buildFill({ type: "none" }));
    }
}
