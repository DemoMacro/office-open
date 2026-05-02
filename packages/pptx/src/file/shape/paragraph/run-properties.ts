import type { ShapeFill } from "@file/drawingml/shape-properties";
import { BuilderElement, NextAttributeComponent, XmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";

let nextHyperlinkId = 1;

export interface IHyperlinkOptions {
    readonly url: string;
    readonly tooltip?: string;
}

export interface IRunPropertiesOptions {
    readonly fontSize?: number;
    readonly bold?: boolean;
    readonly italic?: boolean;
    readonly underline?: "sng" | "dbl" | "none";
    readonly font?: string;
    readonly lang?: string;
    readonly fill?: ShapeFill;
    readonly hyperlink?: IHyperlinkOptions;
}

/**
 * a:rPr — Run properties (font, size, color, etc.).
 */
export class RunProperties extends XmlComponent {
    private readonly hyperlinkKey?: string;
    private readonly hyperlinkUrl?: string;
    private readonly hyperlinkTooltip?: string;

    public static hasProperties(options: IRunPropertiesOptions): boolean {
        return !!(
            options.fontSize ||
            options.bold !== undefined ||
            options.italic !== undefined ||
            options.underline ||
            options.font ||
            options.lang ||
            options.fill ||
            options.hyperlink
        );
    }

    public constructor(options: IRunPropertiesOptions = {}) {
        super("a:rPr");

        const attrs: Record<
            string,
            { readonly key: string; readonly value: string | number | boolean | undefined }
        > = {};
        if (options.fontSize) attrs.sz = { key: "sz", value: options.fontSize * 100 };
        if (options.bold !== undefined) attrs.b = { key: "b", value: options.bold };
        if (options.italic !== undefined) attrs.i = { key: "i", value: options.italic };
        if (options.underline) attrs.u = { key: "u", value: options.underline };
        if (options.lang) attrs.lang = { key: "lang", value: options.lang };

        this.root.push(new NextAttributeComponent(attrs));

        if (options.font) {
            this.root.push(
                new BuilderElement({
                    name: "a:latin",
                    attributes: { typeface: { key: "typeface", value: options.font } },
                }),
            );
            this.root.push(
                new BuilderElement({
                    name: "a:ea",
                    attributes: { typeface: { key: "typeface", value: options.font } },
                }),
            );
        }

        if (options.fill) {
            this.root.push(options.fill);
        }

        if (options.hyperlink) {
            const key = `hlink_${nextHyperlinkId++}`;
            this.hyperlinkKey = key;
            this.hyperlinkUrl = options.hyperlink.url;
            this.hyperlinkTooltip = options.hyperlink.tooltip;

            const hlinkAttrs: Record<string, { readonly key: string; readonly value: string }> = {
                rId: { key: "r:id", value: `{hlink:${key}}` },
            };
            if (options.hyperlink.tooltip) {
                hlinkAttrs.tooltip = { key: "tooltip", value: options.hyperlink.tooltip };
            }
            this.root.push(new BuilderElement({ name: "a:hlinkClick", attributes: hlinkAttrs }));
        }
    }

    public prepForXml(context: IContext): IXmlableObject | undefined {
        if (this.hyperlinkKey) {
            const file = context.fileData as { Hyperlinks?: { addHyperlink(key: string, url: string, tooltip?: string): void } };
            file?.Hyperlinks?.addHyperlink(this.hyperlinkKey, this.hyperlinkUrl!, this.hyperlinkTooltip);
        }
        return super.prepForXml(context);
    }
}
