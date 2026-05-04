import { unzipToMap, readXmlFromZip } from "@office-open/core";
import { parseCoreProperties } from "@office-open/core";

import { createPptxParseContext } from "./context";
import { parseSlide } from "./slide";
import type { PptxDocumentJson } from "./types";

export async function parsePptx(data: Uint8Array): Promise<PptxDocumentJson> {
    const zip = unzipToMap(data);
    const ctx = createPptxParseContext(zip);

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

    return {
        slides,
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
    TableJson,
    TableRowJson,
    TableCellJson,
    ParagraphJson,
    RunJson,
} from "./types";
