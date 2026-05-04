import { uint8ToBase64, getImageType } from "@office-open/core";
import { parseRels, type Relationship } from "@office-open/core";

export interface DocxParseContext {
    zip: Map<string, Uint8Array>;
    hyperlinks: Map<string, string>;
    media: Map<string, { data: string; type: string }>;
    documentRels: Map<string, Relationship>;
}

export function createDocxParseContext(zip: Map<string, Uint8Array>): DocxParseContext {
    const hyperlinks = new Map<string, string>();
    const media = new Map<string, { data: string; type: string }>();
    const documentRels = new Map<string, Relationship>();

    // Parse document.xml.rels
    const rels = parseRels(zip, "word/_rels/document.xml.rels");
    for (const rel of rels) {
        documentRels.set(rel.id, rel);

        // External hyperlinks
        if (rel.targetMode === "External" || rel.type.includes("hyperlink")) {
            hyperlinks.set(rel.id, rel.target);
        }
    }

    // Extract all media files from word/media/
    for (const [path, data] of zip) {
        if (path.startsWith("word/media/")) {
            const fileName = path.split("/").pop() ?? path;
            media.set(path, {
                data: uint8ToBase64(data),
                type: getImageType(fileName),
            });
        }
    }

    return { zip, hyperlinks, media, documentRels };
}
