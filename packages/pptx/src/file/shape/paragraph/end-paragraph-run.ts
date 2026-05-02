import { BuilderElement } from "@file/xml-components";

/**
 * a:endParaRPr — End paragraph run properties.
 */
export class EndParagraphRunProperties extends BuilderElement<{
    readonly lang: string;
}> {
    public constructor(lang: string = "en-US") {
        super({
            name: "a:endParaRPr",
            attributes: {
                lang: { key: "lang", value: lang },
            },
        });
    }
}
