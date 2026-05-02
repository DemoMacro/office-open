/**
 * Core Properties module for PresentationML documents.
 *
 * Reference: ISO-IEC29500-4_2016 shared-documentPropertiesCore.xsd
 *
 * @module
 */
import { StringContainer, XmlAttributeComponent, XmlComponent } from "@file/xml-components";
import { dateTimeValue } from "@util/values";

export interface ICorePropertiesOptions {
    readonly title?: string;
    readonly subject?: string;
    readonly creator?: string;
    readonly keywords?: string;
    readonly description?: string;
    readonly lastModifiedBy?: string;
    readonly revision?: number;
}

export class CoreProperties extends XmlComponent {
    public constructor(options: ICorePropertiesOptions) {
        super("cp:coreProperties");
        this.root.push(new CorePropertiesAttributes(["cp", "dc", "dcterms", "dcmitype", "xsi"]));
        if (options.title) {
            this.root.push(new StringContainer("dc:title", options.title));
        }
        if (options.subject) {
            this.root.push(new StringContainer("dc:subject", options.subject));
        }
        if (options.creator) {
            this.root.push(new StringContainer("dc:creator", options.creator));
        }
        if (options.keywords) {
            this.root.push(new StringContainer("cp:keywords", options.keywords));
        }
        if (options.description) {
            this.root.push(new StringContainer("dc:description", options.description));
        }
        if (options.lastModifiedBy) {
            this.root.push(new StringContainer("cp:lastModifiedBy", options.lastModifiedBy));
        }
        if (options.revision) {
            this.root.push(new StringContainer("cp:revision", String(options.revision)));
        }
        this.root.push(new TimestampElement("dcterms:created"));
        this.root.push(new TimestampElement("dcterms:modified"));
    }
}

/* CSpell:disable */
const PresentationAttributeNamespaces = {
    cp: "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    dc: "http://purl.org/dc/elements/1.1/",
    dcmitype: "http://purl.org/dc/dcmitype/",
    dcterms: "http://purl.org/dc/terms/",
    xsi: "http://www.w3.org/2001/XMLSchema-instance",
};
/* CSpell:enable */

type PresentationAttributeNamespace = keyof typeof PresentationAttributeNamespaces;

class CorePropertiesAttributes extends XmlAttributeComponent<
    Partial<Record<PresentationAttributeNamespace, string>>
> {
    protected readonly xmlKeys: Record<PresentationAttributeNamespace, string> = {
        cp: "xmlns:cp",
        dc: "xmlns:dc",
        dcmitype: "xmlns:dcmitype",
        dcterms: "xmlns:dcterms",
        xsi: "xmlns:xsi",
    };

    public constructor(ns: readonly PresentationAttributeNamespace[]) {
        super(Object.fromEntries(ns.map((n) => [n, PresentationAttributeNamespaces[n]])));
    }
}

class TimestampElementProperties extends XmlAttributeComponent<{ readonly type: string }> {
    protected readonly xmlKeys = { type: "xsi:type" };
}

class TimestampElement extends XmlComponent {
    public constructor(name: string) {
        super(name);
        this.root.push(new TimestampElementProperties({ type: "dcterms:W3CDTF" }));
        this.root.push(dateTimeValue(new Date()));
    }
}
