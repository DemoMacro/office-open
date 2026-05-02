import { XmlAttributeComponent } from "@file/xml-components";

export class AppPropertiesAttributes extends XmlAttributeComponent<{
    readonly xmlns: string;
    readonly vt: string;
}> {
    protected readonly xmlKeys = {
        vt: "xmlns:vt",
        xmlns: "xmlns",
    };
}
