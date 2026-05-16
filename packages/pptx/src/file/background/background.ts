import type { IEffectsOptions } from "@file/drawingml/effects";
import { createPptxEffectList } from "@file/drawingml/effects";
import type { FillOptions } from "@file/drawingml/fill";
import { buildFill } from "@file/drawingml/fill";
import type { IContext, IXmlableObject } from "@file/xml-components";
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
    private readonly shadeToTitle: boolean;

    public constructor(options: IBackgroundOptions) {
        super("p:bgPr");
        this.shadeToTitle = !!options.shadeToTitle;
        this.root.push(
            options.fill !== undefined ? buildFill(options.fill) : buildFill({ type: "none" }),
        );
        if (options.effects) {
            const el = createPptxEffectList(options.effects);
            if (el) this.root.push(el);
        }
    }

    public override prepForXml(context: IContext): IXmlableObject | undefined {
        const obj = super.prepForXml(context);
        if (!obj) return undefined;
        // shadeToTitle is an attribute of p:bgPr, not a child element
        if (this.shadeToTitle && "p:bgPr" in obj) {
            const children = (obj as Record<string, unknown>)["p:bgPr"] as Record<
                string,
                unknown
            >[];
            for (const child of children) {
                if ("_attr" in child) {
                    (child as { _attr: Record<string, unknown> })._attr.shadeToTitle = 1;
                    return obj;
                }
            }
            // No _attr found — insert one
            children.unshift({ _attr: { shadeToTitle: 1 } });
        }
        return obj;
    }
}
