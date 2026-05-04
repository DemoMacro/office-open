import { findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { DocxParseContext } from "./context";
import { parseParagraph } from "./paragraph";
import { parseRun } from "./run";
import { parseTable } from "./table";
import type { ParagraphChildJson, SdtJson, SdtRunJson } from "./types";
import type { FileChildJson } from "./types";

export function parseSdtPr(sdt: Element): Element | undefined {
    return findChild(sdt, "w:sdtPr");
}

/** Parse block-level SDT (w:sdt in body or table cell) */
export function parseSdtContent(sdt: Element, ctx: DocxParseContext): SdtJson | undefined {
    const sdtPr = parseSdtPr(sdt);
    const sdtContent = findChild(sdt, "w:sdtContent");
    if (!sdtContent) {
        return sdtPr ? { $type: "sdt", sdtPr } : undefined;
    }

    const content: FileChildJson[] = [];
    for (const child of sdtContent.elements ?? []) {
        if (child.name === "w:p") {
            content.push(parseParagraph(child, ctx));
        } else if (child.name === "w:tbl") {
            content.push(parseTable(child, ctx));
        } else if (child.name === "w:sdt") {
            const nested = parseSdtContent(child, ctx);
            if (nested) content.push(nested);
        }
    }

    const result: SdtJson = { $type: "sdt" };
    if (sdtPr) result.sdtPr = sdtPr;
    if (content.length > 0) result.content = content;
    return result;
}

/** Parse inline-level SDT (w:sdt inside w:p) */
export function parseSdtRun(sdt: Element, ctx: DocxParseContext): SdtRunJson | undefined {
    const sdtPr = parseSdtPr(sdt);
    const sdtContent = findChild(sdt, "w:sdtContent");
    if (!sdtContent) {
        return sdtPr ? { $type: "sdtRun", sdtPr } : undefined;
    }

    const content: ParagraphChildJson[] = [];
    for (const child of sdtContent.elements ?? []) {
        if (child.name === "w:r") {
            const run = parseRun(child, ctx);
            if (run) content.push(run);
        } else if (child.name === "w:sdt") {
            const nested = parseSdtRun(child, ctx);
            if (nested) content.push(nested);
        }
    }

    const result: SdtRunJson = { $type: "sdtRun" };
    if (sdtPr) result.sdtPr = sdtPr;
    if (content.length > 0) result.content = content;
    return result;
}
