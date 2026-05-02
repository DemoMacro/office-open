import type { IHyperlinkData } from "@file/hyperlink-collection";

export class HyperlinkReplacer {
    public replace(xmlData: string, hyperlinks: readonly IHyperlinkData[], offset: number): string {
        let currentXml = xmlData;

        hyperlinks.forEach((hyperlink, i) => {
            currentXml = currentXml.replace(
                new RegExp(`\\{hlink:${hyperlink.key}\\}`, "g"),
                `rId${offset + i}`,
            );
        });

        return currentXml;
    }
}
