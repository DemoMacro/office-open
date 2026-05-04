/**
 * Sub-document reference module for WordprocessingML documents.
 *
 * SubDoc (w:subDoc) references an external Word document (.docx) that
 * is included as part of the current document. The referenced document
 * is stored as a separate part in the DOCX package.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_Rel (w:subDoc)
 *
 * @module
 */
import type { FileChild } from "@file/file-child";
import { XmlAttributeComponent, XmlComponent } from "@file/xml-components";
import type { IContext, IXmlableObject } from "@file/xml-components";
import { uniqueId } from "@util/convenience-functions";

const SUBDOC_RELATIONSHIP_TYPE =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument";

/**
 * @internal
 */
class SubDocAttributes extends XmlAttributeComponent<{ readonly id: string }> {
    protected readonly xmlKeys = { id: "r:id" };
}

/**
 * Options for creating a SubDoc element.
 */
export interface ISubDocOptions {
    /** The sub-document data (raw bytes of a .docx file) */
    readonly data: Uint8Array | string;
}

/**
 * A sub-document reference element (w:subDoc).
 *
 * References an external Word document that is stored as a separate part
 * in the DOCX package and linked via a relationship.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_Rel
 *
 * @example
 * ```typescript
 * new SubDoc({
 *   data: fs.readFileSync("chapter1.docx"),
 * });
 * ```
 */
export class SubDoc extends XmlComponent implements FileChild {
    public readonly fileChild = Symbol();
    private readonly options: ISubDocOptions;

    public constructor(options: ISubDocOptions) {
        super("w:subDoc");
        this.options = options;
    }

    public prepForXml(context: IContext): IXmlableObject {
        const relId = uniqueId();
        const partPath = `subdocs/subdoc${relId}.docx`;

        this.root.splice(0, 0, new SubDocAttributes({ id: `rId${relId}` }));

        const data =
            typeof this.options.data === "string"
                ? new TextEncoder().encode(this.options.data)
                : this.options.data;

        context.viewWrapper.Relationships.addRelationship(
            relId,
            SUBDOC_RELATIONSHIP_TYPE,
            partPath,
        );

        context.file.SubDocs.addSubDoc(relId, {
            data,
            path: partPath,
        });

        context.file.ContentTypes.addSubDoc(`/word/${partPath}`);

        return super.prepForXml(context) as IXmlableObject;
    }
}
