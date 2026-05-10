import type { Element } from "@office-open/xml";

import type { ParsedDocument } from "@office-open/core";
import { parseDocument } from "@office-open/core";

export { parseDocument };

export interface DocxDocument {
    doc: ParsedDocument;
    body: Element;
}

export function parseDocx(data: Uint8Array): DocxDocument {
    const doc = parseDocument(data);
    const documentEl = doc.get("word/document.xml");
    if (!documentEl) throw new Error("word/document.xml not found");
    const body = documentEl.elements?.find((e) => e.name === "w:body");
    if (!body) throw new Error("w:body not found in word/document.xml");
    return { doc, body };
}
