import type { FillOptions } from "@file/drawingml/fill";
import { buildFill } from "@file/drawingml/fill";
import { XmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";
import { createEffectList, createOutline } from "@office-open/core/drawingml";
import { attrs, xml } from "@office-open/xml";

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
    /** Font size in points. Serialized as OOXML `a:sz` (hundredths of a point). */
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
    effectListObject?: IXmlableObject,
    outlineObject?: IXmlableObject,
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
    if (options.capitalization)
        attrs.cap = TextCapitalization[options.capitalization] ?? options.capitalization;
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

    if (outlineObject) {
        children.push(outlineObject);
    }

    if (effectListObject) {
        children.push(effectListObject);
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

        // Handle outline
        let outlineObj: IXmlableObject | undefined;
        if (opts.outline) {
            outlineObj =
                createOutline({
                    width: 12700,
                    type: "solidFill",
                    color: { value: "000000" },
                }).prepForXml(context) ?? undefined;
        }

        // Handle shadow
        let effectListObj: IXmlableObject | undefined;
        if (opts.shadow) {
            effectListObj =
                createEffectList({
                    outerShadow: {
                        blurRadius: 50800,
                        distance: 38100,
                        direction: 2700000,
                        color: { value: "000000", transforms: { alpha: 40000 } },
                    },
                }).prepForXml(context) ?? undefined;
        }

        return buildRunProperties(opts, hyperlinkKey, fillObj, effectListObj, outlineObj);
    }

    public toXml(context: IContext): string {
        const opts = this.options;

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

        let fillObj: IXmlableObject | undefined;
        if (opts.fill !== undefined) {
            fillObj = buildFill(opts.fill).prepForXml(context) ?? undefined;
        }

        let outlineObj: IXmlableObject | undefined;
        if (opts.outline) {
            outlineObj =
                createOutline({
                    width: 12700,
                    type: "solidFill",
                    color: { value: "000000" },
                }).prepForXml(context) ?? undefined;
        }

        let effectListObj: IXmlableObject | undefined;
        if (opts.shadow) {
            effectListObj =
                createEffectList({
                    outerShadow: {
                        blurRadius: 50800,
                        distance: 38100,
                        direction: 2700000,
                        color: { value: "000000", transforms: { alpha: 40000 } },
                    },
                }).prepForXml(context) ?? undefined;
        }

        return buildRunPropertiesXml(opts, hyperlinkKey, fillObj, effectListObj, outlineObj);
    }
}

/**
 * String version of buildRunProperties for zero-allocation serialization.
 */
export function buildRunPropertiesXml(
    options: IRunPropertiesOptions,
    hyperlinkKey?: string,
    fillObject?: IXmlableObject,
    effectListObject?: IXmlableObject,
    outlineObject?: IXmlableObject,
): string {
    const a: Record<string, string | number | boolean | undefined> = {};
    if (options.fontSize) a.sz = options.fontSize * 100;
    if (options.bold !== undefined) a.b = options.bold;
    if (options.italic !== undefined) a.i = options.italic;
    if (options.underline) a.u = UnderlineStyle[options.underline];
    if (options.lang) a.lang = options.lang;
    if (options.strike) a.strike = StrikeStyle[options.strike];
    if (options.baseline !== undefined) a.baseline = options.baseline;
    if (options.capitalization)
        a.cap = TextCapitalization[options.capitalization] ?? options.capitalization;
    if (options.spacing !== undefined) a.spc = options.spacing;
    if (options.noProof !== undefined) a.noProof = options.noProof;
    if (options.dirty !== undefined) a.dirty = options.dirty;

    const attrStr = attrs(a);
    const body: string[] = [];

    if (options.font) {
        body.push(
            `<a:latin${attrs({ typeface: options.font })}/><a:ea${attrs({ typeface: options.font })}/>`,
        );
    }

    if (options.hyperlink && hyperlinkKey) {
        const h: Record<string, string | undefined> = { "r:id": `{hlink:${hyperlinkKey}}` };
        if (options.hyperlink.tooltip) h.tooltip = options.hyperlink.tooltip;
        body.push(`<a:hlinkClick${attrs(h)}/>`);
    }

    if (options.rightToLeft !== undefined) {
        body.push(`<a:rtl${attrs({ val: options.rightToLeft ? 1 : 0 })}/>`);
    }

    if (fillObject) {
        body.push(xml(fillObject));
    }

    if (outlineObject) {
        body.push(xml(outlineObject));
    }

    if (effectListObject) {
        body.push(xml(effectListObject));
    }

    if (!attrStr && body.length === 0) return "";
    if (body.length === 0) return `<a:rPr${attrStr}/>`;
    return `<a:rPr${attrStr}>${body.join("")}</a:rPr>`;
}
