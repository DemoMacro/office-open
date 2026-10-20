/**
 * AltChunk parser for DOCX documents.
 *
 * Parses w:altChunk elements and extracts embedded content from the ZIP.
 *
 * @module
 */
import { attr } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type { AltChunkOptions } from "@parts/alt-chunk/alt-chunk";

import type { DocxReadContext } from "../../context";

/**
 * Parse a w:altChunk element into AltChunkOptions.
 * Reads the referenced data from the ZIP package.
 */
export function parseAltChunk(el: Element, ctx: DocxReadContext): AltChunkOptions {
  const rId = attr(el, "r:id");
  if (!rId) {
    throw new Error("w:altChunk missing r:id attribute");
  }

  // Look up the path from relationships
  const path = ctx.docx.partRefs.afChunks.get(rId);
  if (!path) {
    throw new Error(`AltChunk relationship ${rId} not found`);
  }

  // Read raw data from ZIP
  const data = ctx.docx.doc.getRaw(path);
  if (!data) {
    throw new Error(`AltChunk data not found at ${path}`);
  }
  // sourcePath is relative to word/ (AltChunkData/compiler convention).
  const sourcePath = path.replace(/^word\//, "");

  // Determine content type from extension
  const ext = sourcePath.split(".").pop() ?? "txt";
  let contentType: AltChunkOptions["contentType"];
  let extension: AltChunkOptions["extension"];

  switch (ext) {
    case "html":
      contentType = "text/html";
      extension = "html";
      break;
    case "rtf":
      contentType = "application/rtf";
      extension = "rtf";
      break;
    case "docx":
      contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
      extension = "docx";
      break;
    default:
      contentType = "text/plain";
      extension = "txt";
      break;
  }

  return {
    data,
    contentType,
    extension,
    sourcePath,
    sourceRid: rId,
  };
}
