/**
 * Core Properties module for PresentationML documents.
 *
 * Reference: ISO-IEC29500-4_2016 shared-documentPropertiesCore.xsd
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";

export interface ICorePropertiesOptions {
    readonly title?: string;
    readonly subject?: string;
    readonly creator?: string;
    readonly keywords?: string;
    readonly description?: string;
    readonly lastModifiedBy?: string;
    readonly revision?: number;
}

export class CoreProperties extends BaseXmlComponent {
    private readonly options: ICorePropertiesOptions;

    public constructor(options: ICorePropertiesOptions) {
        super("cp:coreProperties");
        this.options = options;
    }

    public override prepForXml(_context: IContext): IXmlableObject {
        const opts = this.options;
        const children: IXmlableObject[] = [];

        children.push({
            _attr: {
                "xmlns:cp":
                    "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
                "xmlns:dc": "http://purl.org/dc/elements/1.1/",
                "xmlns:dcmitype": "http://purl.org/dc/dcmitype/",
                "xmlns:dcterms": "http://purl.org/dc/terms/",
                "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
            },
        });

        if (opts.title) children.push({ "dc:title": [opts.title] });
        if (opts.subject) children.push({ "dc:subject": [opts.subject] });
        if (opts.creator) children.push({ "dc:creator": [opts.creator] });
        if (opts.keywords) children.push({ "cp:keywords": [opts.keywords] });
        if (opts.description) children.push({ "dc:description": [opts.description] });
        if (opts.lastModifiedBy) children.push({ "cp:lastModifiedBy": [opts.lastModifiedBy] });
        if (opts.revision) children.push({ "cp:revision": [String(opts.revision)] });

        const now = new Date().toISOString();
        children.push({ "dcterms:created": [{ _attr: { "xsi:type": "dcterms:W3CDTF" } }, now] });
        children.push({ "dcterms:modified": [{ _attr: { "xsi:type": "dcterms:W3CDTF" } }, now] });

        return { "cp:coreProperties": children };
    }
}
