import {
    BuilderElement,
    IgnoreIfEmptyXmlComponent,
    NextAttributeComponent,
} from "@file/xml-components";

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

/**
 * a:pPr — Paragraph properties (alignment, indent, spacing).
 * Omitted from output when completely empty.
 */
export class ParagraphProperties extends IgnoreIfEmptyXmlComponent {
    public constructor(options: IParagraphPropertiesOptions = {}) {
        super("a:pPr");

        const attrs: Record<
            string,
            { readonly key: string; readonly value: string | number | undefined }
        > = {};
        if (options.alignment)
            attrs.algn = { key: "algn", value: TextAlignment[options.alignment] };
        if (options.indentLevel !== undefined)
            attrs.lvl = { key: "lvl", value: options.indentLevel };
        if (options.marginIndent !== undefined)
            attrs.marL = { key: "marL", value: options.marginIndent };
        if (options.marginRight !== undefined)
            attrs.marR = { key: "marR", value: options.marginRight };
        if (options.defTabSize !== undefined)
            attrs.defTabSz = { key: "defTabSz", value: options.defTabSize };
        if (Object.keys(attrs).length > 0) {
            this.root.push(new NextAttributeComponent(attrs));
        }

        if (options.lineSpacing !== undefined) {
            this.root.push(
                new BuilderElement({
                    name: "a:lnSpc",
                    children: [
                        new BuilderElement({
                            name: "a:spcPct",
                            attributes: { val: { key: "val", value: options.lineSpacing * 1000 } },
                        }),
                    ],
                }),
            );
        }

        if (options.lineSpacingPoints !== undefined) {
            this.root.push(
                new BuilderElement({
                    name: "a:lnSpc",
                    children: [
                        new BuilderElement({
                            name: "a:spcPts",
                            attributes: {
                                val: { key: "val", value: options.lineSpacingPoints * 100 },
                            },
                        }),
                    ],
                }),
            );
        }

        if (options.marginBottom !== undefined || options.marginTop !== undefined) {
            this.root.push(
                new BuilderElement({
                    name: "a:spcAft",
                    children: [
                        new BuilderElement({
                            name: "a:spcPts",
                            attributes: { val: { key: "val", value: options.marginBottom ?? 0 } },
                        }),
                    ],
                }),
            );
        }

        if (options.marginTop !== undefined) {
            this.root.push(
                new BuilderElement({
                    name: "a:spcBef",
                    children: [
                        new BuilderElement({
                            name: "a:spcPts",
                            attributes: { val: { key: "val", value: options.marginTop } },
                        }),
                    ],
                }),
            );
        }

        if (options.bullet) {
            this.pushBulletElements(options.bullet);
        } else if (options.bulletNone !== false) {
            this.root.push(new BuilderElement({ name: "a:buNone" }));
        }
    }

    private pushBulletElements(options: BulletOptions) {
        if (options.type === "none") {
            this.root.push(new BuilderElement({ name: "a:buNone" }));
            return;
        }

        // buClr (EG_TextBulletColor)
        if (options.color) {
            this.root.push(
                new BuilderElement({
                    name: "a:buClr",
                    children: [
                        new BuilderElement({
                            name: "a:srgbClr",
                            attributes: {
                                val: { key: "val", value: options.color.replace("#", "") },
                            },
                        }),
                    ],
                }),
            );
        }

        // buSzPct (EG_TextBulletSize) — val is a percentage string like "75%"
        if (options.size !== undefined) {
            this.root.push(
                new BuilderElement({
                    name: "a:buSzPct",
                    attributes: { val: { key: "val", value: `${options.size}%` } },
                }),
            );
        }

        // buFont (EG_TextBulletTypeface)
        this.root.push(
            new BuilderElement({
                name: "a:buFont",
                attributes: {
                    typeface: { key: "typeface", value: "Arial" },
                    panose: { key: "panose", value: "020B0604020202020204" },
                    pitchFamily: { key: "pitchFamily", value: "34" },
                    charset: { key: "charset", value: "0" },
                },
            }),
        );

        // buChar or buAutoNum (EG_TextBullet)
        if (options.type === "char") {
            this.root.push(
                new BuilderElement({
                    name: "a:buChar",
                    attributes: { char: { key: "char", value: options.char ?? "•" } },
                }),
            );
        } else if (options.type === "autoNum") {
            const attrs: Record<string, { readonly key: string; readonly value: string | number }> =
                {
                    type: { key: "type", value: options.format ?? "arabicPeriod" },
                };
            if (options.startAt !== undefined) {
                attrs.startAt = { key: "startAt", value: options.startAt };
            }
            this.root.push(
                new BuilderElement({
                    name: "a:buAutoNum",
                    attributes: attrs,
                }),
            );
        }
    }
}
