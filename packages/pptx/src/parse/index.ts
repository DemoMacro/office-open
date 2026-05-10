import { attr, attrNum, findChild } from "@office-open/xml";

import type { ParsedDocument } from "@office-open/core";
import { parseDocument } from "@office-open/core";

export { parseDocument };

export interface PptxDocument {
    doc: ParsedDocument;
    slidePaths: string[];
    slideWidth?: number;
    slideHeight?: number;
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

    // Parse slide paths from presentation.xml.rels
    const relsXml = doc.get("ppt/_rels/presentation.xml.rels");
    const slidePaths: string[] = [];
    if (relsXml) {
        for (const child of relsXml.elements ?? []) {
            if (child.name !== "Relationship") continue;
            const type = attr(child, "Type") ?? "";
            const target = attr(child, "Target") ?? "";
            if (
                type.includes("/slide") &&
                !type.includes("slideLayout") &&
                !type.includes("slideMaster") &&
                target
            ) {
                const path = target.startsWith("../")
                    ? `ppt/${target.replace("../", "")}`
                    : `ppt/${target}`;
                slidePaths.push(path);
            }
        }
    }
    slidePaths.sort((a, b) => {
        const numA = parseInt(a.match(/slide(\d+)/)?.[1] ?? "0", 10);
        const numB = parseInt(b.match(/slide(\d+)/)?.[1] ?? "0", 10);
        return numA - numB;
    });

    return { doc, slidePaths, slideWidth, slideHeight };
}
