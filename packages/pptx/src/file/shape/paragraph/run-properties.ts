import type { FillOptions } from "@file/drawingml/fill";
import { buildFill } from "@file/drawingml/fill";
import { XmlComponent } from "@file/xml-components";
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
 * Pure function: builds a:rPr XML object from options.
 * @param hyperlinkKey - pre-generated key for hyperlink placeholder
 * @param fillObject - pre-built fill IXmlableObject (from buildFill)
 */
export function buildRunProperties(
    options: IRunPropertiesOptions,
    hyperlinkKey?: string,
    fillObject?: IXmlableObject,
): IXmlableObject | undefined {
    const children: IXmlableObject[] = [];

    const attrs: Record<string, string | number | boolean> = {};
    if (options.fontSize) attrs.sz = options.fontSize * 100;
    if (options.bold !== undefined) attrs.b = options.bold;
    if (options.italic !== undefined) attrs.i = options.italic;
    if (options.underline) attrs.u = UnderlineStyle[options.underline];
    if (options.lang) attrs.lang = options.lang;
    if (options.strike) attrs.strike = StrikeStyle[options.strike];
    if (options.baseline !== undefined) attrs.baseline = options.baseline;
    if (options.capitalization) attrs.cap = TextCapitalization[options.capitalization];
    if (options.spacing !== undefined) attrs.spc = options.spacing;
    if (options.noProof !== undefined) attrs.noProof = options.noProof;
    if (options.dirty !== undefined) attrs.dirty = options.dirty;
    if (Object.keys(attrs).length > 0) children.push({ _attr: attrs });

    if (options.font) {
        children.push({ "a:latin": { _attr: { typeface: options.font } } });
        children.push({ "a:ea": { _attr: { typeface: options.font } } });
    }

    if (options.hyperlink && hyperlinkKey) {
        const hlinkAttrs: Record<string, string> = { "r:id": `{hlink:${hyperlinkKey}}` };
        if (options.hyperlink.tooltip) hlinkAttrs.tooltip = options.hyperlink.tooltip;
        children.push({ "a:hlinkClick": { _attr: hlinkAttrs } });
    }

    if (options.rightToLeft !== undefined) {
        children.push({ "a:rtl": { _attr: { val: options.rightToLeft ? 1 : 0 } } });
    }

    if (fillObject) {
        children.push(fillObject);
    }

    if (children.length === 0) return undefined;

    return {
        "a:rPr": children.length === 1 && "_attr" in children[0] ? children[0] : children,
    };
}

/**
 * a:rPr — Run properties (font, size, color, etc.).
 * Lazy: stores options, builds XML object in prepForXml.
 */
export class RunProperties extends XmlComponent {
    private readonly options: IRunPropertiesOptions;

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
        this.options = options;
    }

    public prepForXml(context: IContext): IXmlableObject | undefined {
        const opts = this.options;

        // Register hyperlinks (B-level: side effect on context)
        let hyperlinkKey: string | undefined;
        if (opts.hyperlink) {
            hyperlinkKey = `hlink_${nextHyperlinkId++}`;
            const file = context.fileData as {
                Hyperlinks?: { addHyperlink(key: string, url: string, tooltip?: string): void };
            };
            file?.Hyperlinks?.addHyperlink(
                hyperlinkKey,
                opts.hyperlink.url,
                opts.hyperlink.tooltip,
            );
        }

        // Handle fill (needs context for prepForXml)
        let fillObj: IXmlableObject | undefined;
        if (opts.fill !== undefined) {
            fillObj = buildFill(opts.fill).prepForXml(context) ?? undefined;
        }

        return buildRunProperties(opts, hyperlinkKey, fillObj);
    }
}
