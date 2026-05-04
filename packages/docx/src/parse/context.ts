import { uint8ToBase64, getImageType } from "@office-open/core";
import { parseRels, type Relationship } from "@office-open/core";

export interface DocxParseContext {
    zip: Map<string, Uint8Array>;
    hyperlinks: Map<string, string>;
    media: Map<string, { data: string; type: string }>;
    documentRels: Map<string, Relationship>;
    mediaPaths: Set<string>;
}

export function createDocxParseContext(zip: Map<string, Uint8Array>): DocxParseContext {
    const hyperlinks = new Map<string, string>();
    const media = new Map<string, { data: string; type: string }>();
    const documentRels = new Map<string, Relationship>();
    const mediaPaths = new Set<string>();

    // Parse document.xml.rels
    const rels = parseRels(zip, "word/_rels/document.xml.rels");
    for (const rel of rels) {
        documentRels.set(rel.id, rel);

        // External hyperlinks
        if (rel.targetMode === "External" || rel.type.includes("hyperlink")) {
            hyperlinks.set(rel.id, rel.target);
        }
    }

    // Collect media paths (lazy conversion)
    for (const path of zip.keys()) {
        if (path.startsWith("word/media/")) {
            mediaPaths.add(path);
        }
    }

    return { zip, hyperlinks, media, documentRels, mediaPaths };
}

export function getMediaData(
    ctx: DocxParseContext,
    path: string,
): { data: string; type: string } | undefined {
    let entry = ctx.media.get(path);
    if (entry) return entry;

    const raw = ctx.zip.get(path);
    if (!raw || !ctx.mediaPaths.has(path)) return undefined;

    const fileName = path.split("/").pop() ?? path;
    entry = { data: uint8ToBase64(raw), type: getImageType(fileName) };
    ctx.media.set(path, entry);
    return entry;
}
