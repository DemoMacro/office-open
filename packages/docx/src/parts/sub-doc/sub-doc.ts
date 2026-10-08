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

import type { DataType } from "@office-open/core";

/**
 * Options for creating a SubDoc element.
 */
export interface SubDocOptions {
  /** The sub-document data: raw .docx bytes, ArrayBuffer, or a base64 data
   *  URL. Required for embedded sub-documents; omit when `sourceUrl` is set. */
  data?: DataType;
  /** Source relationship id in word/_rels/document.xml.rels (round-trip). */
  sourceRid?: string;
  /** External target URL. When set the sub-document is linked, not embedded. */
  sourceUrl?: string;
}
