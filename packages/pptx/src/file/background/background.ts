import { NoFill } from "@file/drawingml/no-fill";
import type { ShapeFill } from "@file/drawingml/shape-properties";
import { XmlComponent } from "@file/xml-components";

export interface IBackgroundOptions {
    readonly fill?: ShapeFill;
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
    public constructor(fill?: ShapeFill) {
        super("p:bgPr");
        this.root.push(fill ?? new NoFill());
    }
}
