import type { IEffectsOptions } from "@file/drawingml/effects";
import { createPptxEffectList } from "@file/drawingml/effects";
import type { FillOptions } from "@file/drawingml/fill";
import { buildFill } from "@file/drawingml/fill";
import { XmlComponent } from "@file/xml-components";

export interface IBackgroundOptions {
    readonly fill?: FillOptions;
    readonly effects?: IEffectsOptions;
    readonly shadeToTitle?: boolean;
}

/**
 * p:bg — Slide background.
 */
export class Background extends XmlComponent {
    public constructor(options: IBackgroundOptions = {}) {
        super("p:bg");
        this.root.push(new BackgroundProperties(options));
    }
}

class BackgroundProperties extends XmlComponent {
    public constructor(options: IBackgroundOptions) {
        super("p:bgPr");
        this.root.push(
            options.fill !== undefined ? buildFill(options.fill) : buildFill({ type: "none" }),
        );
        if (options.effects) {
            const el = createPptxEffectList(options.effects);
            if (el) this.root.push(el);
        }
        if (options.shadeToTitle) {
            this.root.push(
                new (class extends XmlComponent {
                    public constructor() {
                        super("_shadeToTitle");
                    }
                    public override prepForXml() {
                        return { "p:shadeToTitle": {} };
                    }
                })(),
            );
        }
    }
}
