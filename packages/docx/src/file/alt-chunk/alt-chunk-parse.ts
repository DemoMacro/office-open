import type { AltChunkOptions } from "@file/alt-chunk/alt-chunk";
/**
 * AltChunk parser for DOCX documents.
 *
 * Parses w:altChunk elements and extracts embedded content from the ZIP.
 *
 * @module
 */
import { attr } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { ParseContext } from "../../parse/context";

/**
 * Parse a w:altChunk element into AltChunkOptions.
 * Reads the referenced data from the ZIP package.
 */
export function parseAltChunk(el: Element, ctx: ParseContext): AltChunkOptions {
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

  // Determine content type from extension
  const ext = path.split(".").pop() ?? "txt";
  let contentType: "text/html" | "application/rtf" | "text/plain";
  let extension: "html" | "rtf" | "txt";

  switch (ext) {
    case "html":
      contentType = "text/html";
      extension = "html";
      break;
    case "rtf":
      contentType = "application/rtf";
      extension = "rtf";
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
  };
}
