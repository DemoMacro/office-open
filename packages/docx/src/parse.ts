import type { ParsedDocument } from "@office-open/core";
import { parseDocument } from "@office-open/core";
import { attr } from "@office-open/xml";
import type { Element } from "@office-open/xml";

export { parseDocument };

export interface DocxPartRefs {
    /** Header paths keyed by rId */
    headers: Map<string, string>;
    /** Footer paths keyed by rId */
    footers: Map<string, string>;
    /** Path to footnotes.xml if present */
    footnotes?: string;
    /** Path to endnotes.xml if present */
    endnotes?: string;
    /** Path to comments.xml if present */
    comments?: string;
}

export interface DocxDocument {
    doc: ParsedDocument;
    body: Element;
    styles?: Element;
    numbering?: Element;
    settings?: Element;
    fontTable?: Element;
    partRefs: DocxPartRefs;
}

function resolveRelsPath(target: string): string {
    // Target is relative to the source part (word/document.xml), base is "word/"
    if (target.startsWith("/")) return target.slice(1);
    if (target.startsWith("../")) return target.replace("../", "");
    return `word/${target}`;
}

function parsePartRefs(doc: ParsedDocument): DocxPartRefs {
    const relsPath = "word/_rels/document.xml.rels";
    const relsEl = doc.get(relsPath);
    const refs: DocxPartRefs = { headers: new Map(), footers: new Map() };

    if (!relsEl) return refs;

    for (const child of relsEl.elements ?? []) {
        if (child.name !== "Relationship") continue;
        const type = attr(child, "Type") ?? "";
        const target = attr(child, "Target") ?? "";
        const id = attr(child, "Id") ?? "";
        if (!target) continue;

        const path = resolveRelsPath(target);

        if (type.includes("/header")) {
            refs.headers.set(id, path);
        } else if (type.includes("/footer")) {
            refs.footers.set(id, path);
        } else if (type.includes("/footnotes")) {
            refs.footnotes = path;
        } else if (type.includes("/endnotes")) {
            refs.endnotes = path;
        } else if (type.includes("/comments")) {
            refs.comments = path;
        }
    }

    return refs;
}

export function parseDocx(data: Uint8Array): DocxDocument {
    const doc = parseDocument(data);

    const documentEl = doc.get("word/document.xml");
    if (!documentEl) throw new Error("word/document.xml not found");
    const body = documentEl.elements?.find((e) => e.name === "w:body");
    if (!body) throw new Error("w:body not found in word/document.xml");

    // Optional parts — may not exist in all documents
    const styles = doc.get("word/styles.xml");
    const numbering = doc.get("word/numbering.xml");
    const settings = doc.get("word/settings.xml");
    const fontTable = doc.get("word/fontTable.xml");

    const partRefs = parsePartRefs(doc);

    return { doc, body, styles, numbering, settings, fontTable, partRefs };
}
