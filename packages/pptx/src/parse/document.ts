import {
    unzipToMap,
    readXmlFromZip,
    readAllXmlParts,
    parseCoreProperties,
} from "@office-open/core";
import type { Element } from "@office-open/xml";

import { createPptxParseContext } from "./context";
import { parseSlide } from "./slide";
import type { PptxDocumentJson } from "./types";

export interface ParsePptxOptions {
    /** Include raw XML parts in $parts for lossless round-trip (default: true) */
    includeRawParts?: boolean;
}

export async function parsePptx(
    data: Uint8Array,
    options?: ParsePptxOptions,
): Promise<PptxDocumentJson> {
    const zip = unzipToMap(data);
    const ctx = createPptxParseContext(zip);
    const includeRawParts = options?.includeRawParts !== false;

    // Parse core properties
    const coreProps = parseCoreProperties(zip);

    // Parse each slide
    const slides = [];

    for (const slidePath of ctx.slidePaths) {
        const slideXml = readXmlFromZip(zip, slidePath);
        if (!slideXml) continue;

        const slide = parseSlide(slideXml, ctx, slidePath);
        slides.push(slide);
    }

    // Collect raw XML parts for round-trip
    let $parts: Record<string, Element> | undefined;
    if (includeRawParts) {
        $parts = readAllXmlParts(zip, {
            skipPaths: ctx.slidePaths,
        });
    }

    return {
        slides,
        ...($parts && { $parts }),
        ...(ctx.slideWidth !== undefined && { slideWidth: ctx.slideWidth }),
        ...(ctx.slideHeight !== undefined && { slideHeight: ctx.slideHeight }),
        ...(coreProps.title && { title: coreProps.title }),
        ...(coreProps.creator && { creator: coreProps.creator }),
        ...(coreProps.subject && { subject: coreProps.subject }),
        ...(coreProps.keywords && { keywords: coreProps.keywords }),
        ...(coreProps.description && { description: coreProps.description }),
        ...(coreProps.lastModifiedBy && { lastModifiedBy: coreProps.lastModifiedBy }),
    };
}

export type {
    PptxDocumentJson,
    SlideJson,
    SlideChildJson,
    ShapeJson,
    PictureJson,
    ConnectorShapeJson,
    GroupShapeJson,
    TableJson,
    TableRowJson,
    TableCellJson,
    ParagraphJson,
    RunJson,
    RawElement,
    MixedChildren,
} from "./types";
export { isRaw } from "./types";
