import { XmlAttributeComponent } from "@file/xml-components";

export class EndnoteAttributes extends XmlAttributeComponent<{
    readonly type?: string;
    readonly id: number;
}> {
    protected readonly xmlKeys = {
        id: "w:id",
        type: "w:type",
    };
}
