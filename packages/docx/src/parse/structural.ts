import { readXmlFromZip } from "@office-open/core";
import { attr } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { DocxParseContext } from "./context";
import { parseParagraph } from "./paragraph";
import type { FootnoteEntry, CommentEntry, FileChildJson } from "./types";

/**
 * Parse footnotes.xml into FootnoteEntry[].
 */
export function parseFootnotes(zip: Map<string, Uint8Array>): FootnoteEntry[] {
    const xml = readXmlFromZip(zip, "word/footnotes.xml");
    if (!xml) return [];

    const entries: FootnoteEntry[] = [];
    for (const child of xml.elements ?? []) {
        if (child.name === "w:footnote") {
            const id = attr(child, "w:id");
            const type = getFootnoteType(child);
            const children = parseNoteContent(child, zip, "footnote");
            entries.push({ id: id ?? "", ...(type && { type }), ...(children && { children }) });
        }
    }
    return entries;
}

/**
 * Parse endnotes.xml into FootnoteEntry[].
 */
export function parseEndnotes(zip: Map<string, Uint8Array>): FootnoteEntry[] {
    const xml = readXmlFromZip(zip, "word/endnotes.xml");
    if (!xml) return [];

    const entries: FootnoteEntry[] = [];
    for (const child of xml.elements ?? []) {
        if (child.name === "w:endnote") {
            const id = attr(child, "w:id");
            const type = getFootnoteType(child);
            const children = parseNoteContent(child, zip, "endnote");
            entries.push({ id: id ?? "", ...(type && { type }), ...(children && { children }) });
        }
    }
    return entries;
}

/**
 * Parse comments.xml into CommentEntry[].
 */
export function parseComments(zip: Map<string, Uint8Array>): CommentEntry[] {
    const xml = readXmlFromZip(zip, "word/comments.xml");
    if (!xml) return [];

    const entries: CommentEntry[] = [];
    for (const child of xml.elements ?? []) {
        if (child.name === "w:comment") {
            const id = attr(child, "w:id");
            const author = attr(child, "w:author");
            const date = attr(child, "w:date");
            const initials = attr(child, "w:initials");
            const children: FileChildJson[] = [];
            for (const p of child.elements ?? []) {
                if (p.name === "w:p") {
                    children.push(parseParagraph(p, createNoteContext(zip)));
                }
            }
            entries.push({
                id: id ?? "",
                ...(author && { author }),
                ...(date && { date }),
                ...(initials && { initials }),
                ...(children.length > 0 && { children }),
            });
        }
    }
    return entries;
}

export function getFootnoteType(
    note: Element,
): "separator" | "continuationSeparator" | "normal" | undefined {
    const type = attr(note, "w:type");
    if (type === "separator") return "separator";
    if (type === "continuationSeparator") return "continuationSeparator";
    return undefined;
}

export function parseNoteContent(
    note: Element,
    zip: Map<string, Uint8Array>,
    _partType: string,
): FileChildJson[] {
    const ctx = createNoteContext(zip);
    const children: FileChildJson[] = [];
    for (const child of note.elements ?? []) {
        if (child.name === "w:p") {
            children.push(parseParagraph(child, ctx));
        }
    }
    return children;
}

/** Create a minimal parse context for note parsing (no hyperlinks/media) */
export function createNoteContext(zip: Map<string, Uint8Array>): DocxParseContext {
    return {
        zip,
        hyperlinks: new Map(),
        media: new Map(),
        documentRels: new Map(),
        mediaPaths: new Set(),
    };
}
