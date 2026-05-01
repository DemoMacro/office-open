import { XmlAttributeComponent } from "@file/xml-components";

export class EndnotesAttributes extends XmlAttributeComponent<{
    readonly wpc?: string;
    readonly mc?: string;
    readonly o?: string;
    readonly r?: string;
    readonly m?: string;
    readonly v?: string;
    readonly wp14?: string;
    readonly wp?: string;
    readonly w10?: string;
    readonly w?: string;
    readonly w14?: string;
    readonly w15?: string;
    readonly wpg?: string;
    readonly wpi?: string;
    readonly wne?: string;
    readonly wps?: string;
    readonly Ignorable?: string;
}> {
    protected readonly xmlKeys = {
        Ignorable: "mc:Ignorable",
        m: "xmlns:m",
        mc: "xmlns:mc",
        o: "xmlns:o",
        r: "xmlns:r",
        v: "xmlns:v",
        w: "xmlns:w",
        w10: "xmlns:w10",
        w14: "xmlns:w14",
        w15: "xmlns:w15",
        wne: "xmlns:wne",
        wp: "xmlns:wp",
        wp14: "xmlns:wp14",
        wpc: "xmlns:wpc",
        wpg: "xmlns:wpg",
        wpi: "xmlns:wpi",
        wps: "xmlns:wps",
    };
}
