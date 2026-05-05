import { attrs, escapeXml } from "@office-open/xml";
import { XmlComponent } from "@file/xml-components";
import type { IXmlableObject } from "@file/xml-components";

export const TextAlignment = {
    LEFT: "l",
    CENTER: "ctr",
    RIGHT: "r",
    JUSTIFY: "just",
} as const;

export type BulletCharOptions = {
    readonly type: "char";
    readonly char?: string;
    readonly color?: string;
    readonly size?: number;
};

export type BulletAutoNumOptions = {
    readonly type: "autoNum";
    readonly format?: string;
    readonly startAt?: number;
    readonly color?: string;
    readonly size?: number;
};

export type BulletNoneOption = {
    readonly type: "none";
};

export type BulletOptions = BulletCharOptions | BulletAutoNumOptions | BulletNoneOption;

export interface IParagraphPropertiesOptions {
    readonly alignment?: keyof typeof TextAlignment;
    readonly indentLevel?: number;
    readonly marginBottom?: number;
    readonly marginTop?: number;
    readonly bullet?: BulletOptions;
    /** @deprecated Use bullet: { type: "none" } instead */
    readonly bulletNone?: boolean;
    readonly lineSpacing?: number;
    readonly lineSpacingPoints?: number;
    readonly marginIndent?: number;
    readonly marginRight?: number;
    readonly defTabSize?: number;
}

function buildBulletChildren(options: BulletOptions): IXmlableObject[] {
    const children: IXmlableObject[] = [];

    if (options.type === "none") {
        children.push({ "a:buNone": {} });
        return children;
    }

    if (options.color) {
        children.push({
            "a:buClr": [{ "a:srgbClr": { _attr: { val: options.color.replace("#", "") } } }],
        });
    }

    if (options.size !== undefined) {
        children.push({ "a:buSzPct": { _attr: { val: `${options.size}%` } } });
    }

    children.push({
        "a:buFont": {
            _attr: {
                typeface: "Arial",
                panose: "020B0604020202020204",
                pitchFamily: "34",
                charset: "0",
            },
        },
    });

    if (options.type === "char") {
        children.push({ "a:buChar": { _attr: { char: options.char ?? "•" } } });
    } else if (options.type === "autoNum") {
        const buAttrs: Record<string, string | number> = { type: options.format ?? "arabicPeriod" };
        if (options.startAt !== undefined) buAttrs.startAt = options.startAt;
        children.push({ "a:buAutoNum": { _attr: buAttrs } });
    }

    return children;
}

/**
 * Pure function: builds a:pPr XML object from options.
 * Returns undefined when no meaningful content (empty → omitted from output).
 */
export function buildParagraphProperties(
    options: IParagraphPropertiesOptions,
): IXmlableObject | undefined {
    const children: IXmlableObject[] = [];

    const attrs: Record<string, string | number> = {};
    if (options.alignment) attrs.algn = TextAlignment[options.alignment];
    if (options.indentLevel !== undefined) attrs.lvl = options.indentLevel;
    if (options.marginIndent !== undefined) attrs.marL = options.marginIndent;
    if (options.marginRight !== undefined) attrs.marR = options.marginRight;
    if (options.defTabSize !== undefined) attrs.defTabSz = options.defTabSize;
    if (Object.keys(attrs).length > 0) children.push({ _attr: attrs });

    if (options.lineSpacing !== undefined) {
        children.push({
            "a:lnSpc": [{ "a:spcPct": { _attr: { val: options.lineSpacing * 1000 } } }],
        });
    }

    if (options.lineSpacingPoints !== undefined) {
        children.push({
            "a:lnSpc": [{ "a:spcPts": { _attr: { val: options.lineSpacingPoints * 100 } } }],
        });
    }

    if (options.marginBottom !== undefined || options.marginTop !== undefined) {
        children.push({
            "a:spcAft": [{ "a:spcPts": { _attr: { val: options.marginBottom ?? 0 } } }],
        });
    }

    if (options.marginTop !== undefined) {
        children.push({
            "a:spcBef": [{ "a:spcPts": { _attr: { val: options.marginTop } } }],
        });
    }

    if (options.bullet) {
        children.push(...buildBulletChildren(options.bullet));
    } else if (options.bulletNone !== false) {
        children.push({ "a:buNone": {} });
    }

    if (children.length === 0) return undefined;

    return {
        "a:pPr": children.length === 1 && "_attr" in children[0] ? children[0] : children,
    };
}

/**
 * a:pPr — Paragraph properties (alignment, indent, spacing).
 * Lazy: stores options, builds XML object in prepForXml.
 * Omitted from output when completely empty.
 */
export class ParagraphProperties extends XmlComponent {
    private readonly options: IParagraphPropertiesOptions;

    public constructor(options: IParagraphPropertiesOptions = {}) {
        super("a:pPr");
        this.options = options;
    }

    public prepForXml(
        _context: import("@file/xml-components").IContext,
    ): IXmlableObject | undefined {
        return buildParagraphProperties(this.options);
    }

    public toXml(): string {
        return buildParagraphPropertiesXml(this.options);
    }
}

function buildBulletChildrenXml(options: BulletOptions): string {
    if (options.type === "none") {
        return "<a:buNone/>";
    }

    let s = "";
    if (options.color) {
        s += `<a:buClr><a:srgbClr${attrs({ val: options.color.replace("#", "") })}/></a:buClr>`;
    }
    if (options.size !== undefined) {
        s += `<a:buSzPct${attrs({ val: `${options.size}%` })}/>`;
    }
    s += `<a:buFont${attrs({ typeface: "Arial", panose: "020B0604020202020204", pitchFamily: "34", charset: "0" })}/>`;
    if (options.type === "char") {
        s += `<a:buChar${attrs({ char: options.char ?? "•" })}/>`;
    } else if (options.type === "autoNum") {
        const buAttrs: Record<string, string | number | undefined> = { type: options.format ?? "arabicPeriod" };
        if (options.startAt !== undefined) buAttrs.startAt = options.startAt;
        s += `<a:buAutoNum${attrs(buAttrs)}/>`;
    }
    return s;
}

export function buildParagraphPropertiesXml(options: IParagraphPropertiesOptions): string {
    const attrStr = attrs({
        algn: options.alignment ? TextAlignment[options.alignment] : undefined,
        lvl: options.indentLevel,
        marL: options.marginIndent,
        marR: options.marginRight,
        defTabSz: options.defTabSize,
    });

    const body: string[] = [];

    if (options.lineSpacing !== undefined) {
        body.push(`<a:lnSpc><a:spcPct${attrs({ val: options.lineSpacing * 1000 })}/></a:lnSpc>`);
    }

    if (options.lineSpacingPoints !== undefined) {
        body.push(`<a:lnSpc><a:spcPts${attrs({ val: options.lineSpacingPoints * 100 })}/></a:lnSpc>`);
    }

    if (options.marginBottom !== undefined || options.marginTop !== undefined) {
        body.push(`<a:spcAft><a:spcPts${attrs({ val: options.marginBottom ?? 0 })}/></a:spcAft>`);
    }

    if (options.marginTop !== undefined) {
        body.push(`<a:spcBef><a:spcPts${attrs({ val: options.marginTop })}/></a:spcBef>`);
    }

    if (options.bullet) {
        body.push(buildBulletChildrenXml(options.bullet));
    } else if (options.bulletNone !== false) {
        body.push("<a:buNone/>");
    }

    if (!attrStr && body.length === 0) return "";
    if (body.length === 0) return `<a:pPr${attrStr}/>`;
    return `<a:pPr${attrStr}>${body.join("")}</a:pPr>`;
}
