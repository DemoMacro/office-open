import { XmlComponent, XmlAttributeComponent } from "../xml-components";

class AppPropertiesAttributes extends XmlAttributeComponent<{
    readonly xmlns: string;
    readonly vt: string;
}> {
    protected readonly xmlKeys = {
        vt: "xmlns:vt",
        xmlns: "xmlns",
    };
}

export class AppProperties extends XmlComponent {
    public constructor() {
        super("Properties");
        this.root.push(
            new AppPropertiesAttributes({
                vt: "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
                xmlns: "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
            }),
        );
    }
}
