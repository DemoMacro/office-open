import type { FillOptions } from "@file/drawingml/fill";
import { buildFill } from "@file/drawingml/fill";
import { BuilderElement, NextAttributeComponent, XmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";

let nextHyperlinkId = 1;

export const UnderlineStyle = {
    SINGLE: "sng",
    DOUBLE: "dbl",
    NONE: "none",
} as const;

export const StrikeStyle = {
    SINGLE: "sngStrike",
    DOUBLE: "dblStrike",
    NONE: "noStrike",
} as const;

export const TextCapitalization = {
    NONE: "none",
    ALL: "all",
    SMALL: "small",
} as const;

export interface IHyperlinkOptions {
    readonly url: string;
    readonly tooltip?: string;
}

export interface IRunPropertiesOptions {
    readonly fontSize?: number;
    readonly bold?: boolean;
    readonly italic?: boolean;
    readonly underline?: keyof typeof UnderlineStyle;
    readonly font?: string;
    readonly lang?: string;
    readonly fill?: FillOptions;
    readonly hyperlink?: IHyperlinkOptions;
    readonly strike?: keyof typeof StrikeStyle;
    readonly baseline?: number;
    readonly spacing?: number;
    readonly capitalization?: keyof typeof TextCapitalization;
    readonly shadow?: boolean;
    readonly outline?: boolean;
    readonly rightToLeft?: boolean;
    readonly noProof?: boolean;
    readonly dirty?: boolean;
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
            options.hyperlink ||
            options.strike ||
            options.baseline !== undefined ||
            options.spacing !== undefined ||
            options.capitalization ||
            options.shadow !== undefined ||
            options.outline !== undefined ||
            options.rightToLeft !== undefined ||
            options.noProof !== undefined ||
            options.dirty !== undefined
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
        if (options.underline) attrs.u = { key: "u", value: UnderlineStyle[options.underline] };
        if (options.lang) attrs.lang = { key: "lang", value: options.lang };
        if (options.strike) attrs.strike = { key: "strike", value: StrikeStyle[options.strike] };
        if (options.baseline !== undefined)
            attrs.baseline = { key: "baseline", value: options.baseline };
        if (options.capitalization)
            attrs.cap = { key: "cap", value: TextCapitalization[options.capitalization] };
        if (options.spacing !== undefined) attrs.spc = { key: "spc", value: options.spacing };
        if (options.noProof !== undefined)
            attrs.noProof = { key: "noProof", value: options.noProof };
        if (options.dirty !== undefined) attrs.dirty = { key: "dirty", value: options.dirty };

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

        if (options.fill !== undefined) {
            this.root.push(buildFill(options.fill));
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

        if (options.rightToLeft !== undefined) {
            this.root.push(
                new BuilderElement({
                    name: "a:rtl",
                    attributes: { val: { key: "val", value: options.rightToLeft ? 1 : 0 } },
                }),
            );
        }
    }

    public prepForXml(context: IContext): IXmlableObject | undefined {
        if (this.hyperlinkKey) {
            const file = context.fileData as {
                Hyperlinks?: { addHyperlink(key: string, url: string, tooltip?: string): void };
            };
            file?.Hyperlinks?.addHyperlink(
                this.hyperlinkKey,
                this.hyperlinkUrl!,
                this.hyperlinkTooltip,
            );
        }
        return super.prepForXml(context);
    }
}
