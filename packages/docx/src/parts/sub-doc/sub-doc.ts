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

/**
 * Options for creating a SubDoc element.
 */
export interface SubDocOptions {
  /** The sub-document data (raw bytes of a .docx file) */
  data: Uint8Array | string;
}
