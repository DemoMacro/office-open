import { FileChild } from "@file/file-child";
import { EmptyElement, XmlAttributeComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";
import { uniqueId } from "@util/convenience-functions";
/**
 * Alternative format chunk module for WordprocessingML documents.
 *
 * AltChunk (w:altChunk) embeds content from alternative formats (HTML, RTF, plain text)
 * into a Word document. The content is stored as a separate part in the DOCX package
 * and referenced via a relationship.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_AltChunk
 *
 * @module
 */

/**
 * @internal
 */
class AltChunkAttributes extends XmlAttributeComponent<{ readonly id: string }> {
    protected readonly xmlKeys = { id: "r:id" };
}

/**
 * Options for creating an AltChunk element.
 */
export interface IAltChunkOptions {
    /** Content data to embed (string or binary) */
    readonly data: Uint8Array | string;
    /** MIME content type of the data */
    readonly contentType: "text/html" | "application/rtf" | "text/plain";
    /** File extension for the part */
    readonly extension: "html" | "rtf" | "txt";
    /** Whether to match source formatting */
    readonly matchSrc?: boolean;
}

const ALTCHUNK_RELATIONSHIP_TYPE =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk";

/**
 * An alternative format chunk element (CT_AltChunk).
 *
 * Embeds content from an alternative format (HTML, RTF, plain text) into
 * a Word document. The content is stored as a separate part and referenced
 * via an `r:id` relationship.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_AltChunk
 *
 * @example
 * ```typescript
 * new AltChunk({
 *   data: "<p>Hello HTML</p>",
 *   contentType: "text/html",
 *   extension: "html",
 * });
 * ```
 */
export class AltChunk extends FileChild {
    private readonly options: IAltChunkOptions;

    public constructor(options: IAltChunkOptions) {
        super("w:altChunk");
        this.options = options;

        if (options.matchSrc) {
            const altChunkPr = new EmptyElement("w:altChunkPr");
            altChunkPr.addChildElement(new EmptyElement("w:matchSrc"));
            this.root.push(altChunkPr);
        }
    }

    public prepForXml(context: IContext): IXmlableObject {
        const relId = uniqueId();
        const extension = this.options.extension;

        this.root.splice(0, 0, new AltChunkAttributes({ id: `rId${relId}` }));

        const partPath = `afchunks/afchunk${relId}.${extension}`;
        const rawData =
            typeof this.options.data === "string"
                ? new TextEncoder().encode(this.options.data)
                : this.options.data;
        const data =
            this.options.contentType === "text/html" && typeof this.options.data === "string"
                ? new TextEncoder().encode(wrapHtmlDocument(this.options.data))
                : rawData;

        context.viewWrapper.Relationships.addRelationship(
            relId,
            ALTCHUNK_RELATIONSHIP_TYPE,
            partPath,
        );

        context.file.AltChunks.addAltChunk(relId, {
            key: relId,
            data,
            path: partPath,
            extension,
            contentType: this.options.contentType,
        });

        context.file.ContentTypes.addAltChunk(
            `/word/${partPath}`,
            this.options.contentType,
            extension,
        );

        return super.prepForXml(context) as IXmlableObject;
    }
}

function wrapHtmlDocument(fragment: string): string {
    if (/<(!DOCTYPE|html|HTML)/i.test(fragment)) {
        return fragment;
    }
    return `<!DOCTYPE html>\n<html><head><meta charset="utf-8"></head>\n<body>${fragment}</body></html>`;
}
