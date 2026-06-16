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
 * Options for creating an AltChunk element.
 */
export interface AltChunkOptions {
  /** Content data to embed (string or binary) */
  data: Uint8Array | string;
  /** MIME content type of the data */
  contentType: "text/html" | "application/rtf" | "text/plain";
  /** File extension for the part */
  extension: "html" | "rtf" | "txt";
  /** Whether to match source formatting (w:matchSrc) */
  matchSource?: boolean;
}
