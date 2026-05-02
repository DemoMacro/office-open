import {
    BuilderElement,
    IgnoreIfEmptyXmlComponent,
    NextAttributeComponent,
} from "@file/xml-components";

export interface IParagraphPropertiesOptions {
    readonly alignment?: "l" | "ctr" | "r" | "just";
    readonly indentLevel?: number;
    readonly marginBottom?: number;
    readonly marginTop?: number;
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
        if (options.alignment) attrs.algn = { key: "algn", value: options.alignment };
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
                            attributes: { val: { key: "val", value: options.lineSpacingPoints * 100 } },
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

        if (options.bulletNone !== false) {
            this.root.push(
                new BuilderElement({
                    name: "a:buNone",
                }),
            );
        }
    }
}
