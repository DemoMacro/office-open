import { ImportedXmlComponent } from "@file/xml-components";

export interface IShowOptions {
    readonly loop?: boolean;
    readonly kiosk?: boolean;
    readonly showNarration?: boolean;
    readonly useTimings?: boolean;
}

function buildPresPropsXml(showOptions?: IShowOptions): string {
    const showPrXml = showOptions
        ? `<p:showPr${[
              showOptions.loop ? ' loop="1"' : "",
              showOptions.kiosk ? ' kiosk="1"' : "",
              showOptions.showNarration === false ? ' showNarration="0"' : "",
              showOptions.useTimings ? ' useTimings="1"' : "",
          ].join("")}><p:present/></p:showPr>`
        : "";
    return `<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">${showPrXml}</p:presentationPr>`;
}

export class PresentationProperties extends ImportedXmlComponent {
    private static cache = new Map<string, ImportedXmlComponent>();

    public constructor(showOptions?: IShowOptions) {
        super("p:presentationPr");
        const key = showOptions ? JSON.stringify(showOptions) : "";
        if (!PresentationProperties.cache.has(key)) {
            PresentationProperties.cache.set(
                key,
                ImportedXmlComponent.fromXmlString(buildPresPropsXml(showOptions)),
            );
        }
    }

    public prepForXml() {
        const key = this.showOptions ? JSON.stringify(this.showOptions) : "";
        return PresentationProperties.cache.get(key)!.prepForXml({ stack: [] });
    }

    private readonly showOptions?: IShowOptions;
}
