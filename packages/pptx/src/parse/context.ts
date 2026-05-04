import { readXmlFromZip, parseRels } from "@office-open/core";
import { findChild, attrNum } from "@office-open/xml";

export interface PptxParseContext {
    zip: Map<string, Uint8Array>;
    slidePaths: string[];
    slideWidth?: number;
    slideHeight?: number;
}

export function createPptxParseContext(zip: Map<string, Uint8Array>): PptxParseContext {
    const slidePaths: string[] = [];
    let slideWidth: number | undefined;
    let slideHeight: number | undefined;

    // Parse presentation.xml
    const presentationXml = readXmlFromZip(zip, "ppt/presentation.xml");
    if (presentationXml) {
        const sldSz = findChild(presentationXml, "p:sldSz");
        if (sldSz) {
            slideWidth = attrNum(sldSz, "cx");
            slideHeight = attrNum(sldSz, "cy");
        }
    }

    // Parse presentation.xml.rels for slide paths
    const rels = parseRels(zip, "ppt/_rels/presentation.xml.rels");
    for (const rel of rels) {
        if (
            rel.type.includes("/slide") &&
            !rel.type.includes("slideLayout") &&
            !rel.type.includes("slideMaster")
        ) {
            const path = rel.target.startsWith("../")
                ? `ppt/${rel.target.replace("../", "")}`
                : `ppt/${rel.target}`;
            slidePaths.push(path);
        }
    }

    // Sort slide paths by number for consistent ordering
    slidePaths.sort((a, b) => {
        const numA = parseInt(a.match(/slide(\d+)/)?.[1] ?? "0", 10);
        const numB = parseInt(b.match(/slide(\d+)/)?.[1] ?? "0", 10);
        return numA - numB;
    });

    return { zip, slidePaths, slideWidth, slideHeight };
}

export function resolveSlideMediaPath(relsTarget: string): string {
    if (relsTarget.startsWith("../")) {
        return `ppt/${relsTarget.replace("../", "")}`;
    }
    return `ppt/${relsTarget}`;
}

export function getSlideRels(
    ctx: PptxParseContext,
    slidePath: string,
): Map<string, { target: string; type: string }> {
    const relsMap = new Map<string, { target: string; type: string }>();

    // Convert slide path to rels path
    const relsPath = slidePath.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
    const rels = parseRels(ctx.zip, relsPath);
    for (const rel of rels) {
        relsMap.set(rel.id, { target: rel.target, type: rel.type });
    }

    return relsMap;
}
