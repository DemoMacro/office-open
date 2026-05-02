import { BuilderElement, XmlComponent } from "@file/xml-components";

export interface IHeaderFooterOptions {
    readonly slideNumber?: boolean;
    readonly dateTime?: boolean;
    readonly footer?: string;
}

/**
 * p:hf — Slide header/footer settings.
 */
export class HeaderFooter extends XmlComponent {
    public constructor(options: IHeaderFooterOptions = {}) {
        super("p:hf");

        if (options.slideNumber !== false) {
            this.root.push(
                new BuilderElement({ name: "p:sldNum" }),
            );
        }

        if (options.dateTime !== false) {
            this.root.push(
                new BuilderElement({ name: "p:dt" }),
            );
        }

        if (options.footer) {
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
                                            children: [
                                                new BuilderElement({
                                                    name: "a:t",
                                                    children: [options.footer],
                                                }),
                                            ],
                                        }),
                                        new BuilderElement({ name: "a:endParaRPr", attributes: { lang: { key: "lang", value: "en-US" } } }),
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
