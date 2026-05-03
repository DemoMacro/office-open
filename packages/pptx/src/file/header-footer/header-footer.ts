import {
    BuilderElement,
    NextAttributeComponent,
    StringContainer,
    XmlComponent,
} from "@file/xml-components";

export interface IHeaderFooterOptions {
    readonly slideNumber?: boolean;
    readonly dateTime?: boolean;
    readonly footer?: string | boolean;
    readonly header?: boolean;
}

/**
 * p:hf — Slide header/footer settings.
 *
 * CT_HeaderFooter has boolean attributes (sldNum, hdr, ftr, dt) for visibility.
 * Footer text content is passed as p:ftr child element (legacy support).
 */
export class HeaderFooter extends XmlComponent {
    public constructor(options: IHeaderFooterOptions = {}) {
        super("p:hf");

        // Boolean visibility attributes
        const attrs: Record<string, { readonly key: string; readonly value: number }> = {};
        if (options.slideNumber !== false) attrs.sldNum = { key: "sldNum", value: 1 };
        if (options.dateTime !== false) attrs.dt = { key: "dt", value: 1 };
        // Only set hdr when explicitly true (no header placeholder in default slide master)
        if (options.header === true) attrs.hdr = { key: "hdr", value: 1 };
        // Only set ftr when explicitly true or footer has string content
        if (options.footer !== false && options.footer !== undefined)
            attrs.ftr = { key: "ftr", value: 1 };

        if (Object.keys(attrs).length > 0) {
            this.root.push(new NextAttributeComponent(attrs));
        }

        // Footer text content as child element
        if (typeof options.footer === "string") {
            this.root.push(
                new BuilderElement({
                    name: "p:ftr",
                    children: [
                        new BuilderElement({
                            name: "p:txBody",
                            children: [
                                new BuilderElement({ name: "a:bodyPr" }),
                                new BuilderElement({ name: "a:lstStyle" }),
                                new BuilderElement({
                                    name: "a:p",
                                    children: [
                                        new BuilderElement({
                                            name: "a:r",
                                            children: [new StringContainer("a:t", options.footer)],
                                        }),
                                        new BuilderElement({
                                            name: "a:endParaRPr",
                                            attributes: { lang: { key: "lang", value: "en-US" } },
                                        }),
                                    ],
                                }),
                            ],
                        }),
                    ],
                }),
            );
        }
    }
}
