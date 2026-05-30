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
import { XmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { uniqueId } from "@util/convenience-functions";

const SUBDOC_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument";

/**
 * Options for creating a SubDoc element.
 */
export interface SubDocOptions {
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
  private readonly options: SubDocOptions;

  public constructor(options: SubDocOptions) {
    super("w:subDoc");
    this.options = options;
  }

  public override toXml(context: Context): string {
    const relId = uniqueId();
    const partPath = `subdocs/subdoc${relId}.docx`;

    const data =
      typeof this.options.data === "string"
        ? new TextEncoder().encode(this.options.data)
        : this.options.data;

    context.viewWrapper.relationships.addRelationship(relId, SUBDOC_RELATIONSHIP_TYPE, partPath);

    context.file.subDocs.addSubDoc(relId, {
      data,
      path: partPath,
    });

    context.file.contentTypes.addSubDoc(`/word/${partPath}`);

    this.root.splice(0, 0, { _attr: { "r:id": `rId${relId}` } });
    const result = super.toXml(context);
    this.root.splice(0, 1);
    return result;
  }
}
