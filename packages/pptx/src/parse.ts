import type { ParsedDocument } from "@office-open/core";
import { parseDocument } from "@office-open/core";
import { attr, attrNum, findChild } from "@office-open/xml";

export { parseDocument };

export interface PptxDocument {
    doc: ParsedDocument;
    slidePaths: string[];
    slideMasterPaths: string[];
    slideLayoutPaths: string[];
    notesSlidePaths: string[];
    slideWidth?: number;
    slideHeight?: number;
}

function resolveRelsPath(target: string): string {
    // Target is relative to the source part (ppt/presentation.xml), base is "ppt/"
    if (target.startsWith("/")) return target.slice(1);
    if (target.startsWith("../")) return target.replace("../", "");
    return `ppt/${target}`;
}

function sortByNumber(paths: string[]): string[] {
    return paths.sort((a, b) => {
        const numA = parseInt(a.match(/(\d+)/)?.[1] ?? "0", 10);
        const numB = parseInt(b.match(/(\d+)/)?.[1] ?? "0", 10);
        return numA - numB;
    });
}

/** Filter keys to XML files only (exclude .rels and binary files). */
function xmlKeys(keys: string[]): string[] {
    return keys.filter((k) => k.endsWith(".xml"));
}

export function parsePptx(data: Uint8Array): PptxDocument {
    const doc = parseDocument(data);

    // Parse slide size from presentation.xml
    const presentationXml = doc.get("ppt/presentation.xml");
    let slideWidth: number | undefined;
    let slideHeight: number | undefined;
    if (presentationXml) {
        const sldSz = findChild(presentationXml, "p:sldSz");
        if (sldSz) {
            slideWidth = attrNum(sldSz, "cx");
            slideHeight = attrNum(sldSz, "cy");
        }
    }

    // Parse slide/slideMaster/notesSlide paths from presentation.xml.rels
    const relsXml = doc.get("ppt/_rels/presentation.xml.rels");
    const slidePaths: string[] = [];
    const slideMasterPaths: string[] = [];

    if (relsXml) {
        for (const child of relsXml.elements ?? []) {
            if (child.name !== "Relationship") continue;
            const type = attr(child, "Type") ?? "";
            const target = attr(child, "Target") ?? "";
            if (!target) continue;

            const path = resolveRelsPath(target);

            if (type.includes("/slideMaster")) {
                slideMasterPaths.push(path);
            } else if (
                type.includes("/slide") &&
                !type.includes("slideLayout") &&
                !type.includes("slideMaster")
            ) {
                slidePaths.push(path);
            }
        }
    }

    sortByNumber(slidePaths);
    sortByNumber(slideMasterPaths);

    // slideLayouts are referenced from slideMaster rels, not presentation.xml.rels
    const slideLayoutPaths = sortByNumber(xmlKeys(doc.keys("ppt/slideLayouts/")));

    // notesSlides are referenced from slide rels, not presentation.xml.rels
    const notesSlidePaths = sortByNumber(xmlKeys(doc.keys("ppt/notesSlides/")));

    return {
        doc,
        slidePaths,
        slideMasterPaths,
        slideLayoutPaths,
        notesSlidePaths,
        slideWidth,
        slideHeight,
    };
}
