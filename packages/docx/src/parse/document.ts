import {
    unzipToMap,
    readXmlFromZip,
    readAllXmlParts,
    parseCoreProperties,
} from "@office-open/core";
import { findChild, type Element } from "@office-open/xml";

import { createDocxParseContext, type DocxParseContext } from "./context";
import { parseNumbering, buildNumberingConfig } from "./numbering";
import { parseParagraph } from "./paragraph";
import { parseSdtContent } from "./sdt";
import { parseSectionProperties } from "./section";
import { parseFootnotes, parseEndnotes, parseComments } from "./structural";
import { parseTable } from "./table";
import type { DocxDocumentJson, SectionJson, FileChildJson } from "./types";

export interface ParseDocxOptions {
    /** Include raw XML parts in $parts for lossless round-trip (default: true) */
    includeRawParts?: boolean;
}

export async function parseDocx(
    data: Uint8Array,
    options?: ParseDocxOptions,
): Promise<DocxDocumentJson> {
    const zip = unzipToMap(data);
    const ctx = createDocxParseContext(zip);
    const includeRawParts = options?.includeRawParts !== false;

    // Parse core properties
    const coreProps = parseCoreProperties(zip);

    // Parse document.xml
    const documentXml = readXmlFromZip(zip, "word/document.xml");
    if (!documentXml) {
        throw new Error("Invalid DOCX file: missing word/document.xml");
    }

    const body = findChild(documentXml, "w:body");
    if (!body) {
        throw new Error("Invalid DOCX file: missing w:body element");
    }

    // Parse body content, splitting at section breaks
    const sections = parseBodySections(body, ctx);

    // Resolve header/footer content from relationship IDs
    resolveHeadersFooters(sections, ctx);

    // Parse numbering.xml and build config
    const numberingData = parseNumbering(zip);
    const numberingConfig = buildNumberingConfig(numberingData, sections);

    // Parse structural parts
    const footnotes = parseFootnotes(zip);
    const endnotes = parseEndnotes(zip);
    const comments = parseComments(zip);

    // Collect raw XML parts for round-trip
    let $parts: Record<string, Element> | undefined;
    if (includeRawParts) {
        $parts = readAllXmlParts(zip, {
            skipPaths: ["word/document.xml"],
        });
    }

    return {
        sections,
        ...($parts && { $parts }),
        ...(numberingConfig.length > 0 && { numbering: numberingConfig }),
        ...(footnotes.length > 0 && { footnotes }),
        ...(endnotes.length > 0 && { endnotes }),
        ...(comments.length > 0 && { comments }),
        ...(coreProps.title && { title: coreProps.title }),
        ...(coreProps.subject && { subject: coreProps.subject }),
        ...(coreProps.creator && { creator: coreProps.creator }),
        ...(coreProps.keywords && { keywords: coreProps.keywords }),
        ...(coreProps.description && { description: coreProps.description }),
        ...(coreProps.lastModifiedBy && { lastModifiedBy: coreProps.lastModifiedBy }),
        ...(coreProps.revision && { revision: coreProps.revision }),
    };
}

function parseBodySections(
    body: Element,
    ctx: ReturnType<typeof createDocxParseContext>,
): SectionJson[] {
    const sections: SectionJson[] = [];
    let currentChildren: FileChildJson[] = [];

    // Collect all body elements except sectPr (which is at the end)
    const bodyElements = (body.elements ?? []).filter((el: Element) => el.name !== "w:sectPr");

    // Find the final sectPr (document-level section properties)
    const finalSectPr = findChild(body, "w:sectPr");

    // Process all body elements
    for (const element of bodyElements) {
        if (element.name === "w:p") {
            // Check for section break (sectPr inside pPr)
            const pPr = findChild(element, "w:pPr");
            if (pPr) {
                const sectPr = findChild(pPr, "w:sectPr");
                if (sectPr) {
                    // Section break found — finalize current section
                    const paragraph = parseParagraph(element, ctx);
                    // Remove section break properties from the paragraph
                    currentChildren.push(paragraph);

                    sections.push({
                        properties: parseSectionProperties(sectPr),
                        children: currentChildren,
                    });
                    currentChildren = [];
                    continue;
                }
            }
            currentChildren.push(parseParagraph(element, ctx));
        } else if (element.name === "w:tbl") {
            currentChildren.push(parseTable(element, ctx));
        } else if (element.name === "w:sdt") {
            const sdt = parseSdtContent(element, ctx);
            if (sdt) currentChildren.push(sdt);
            else currentChildren.push({ $raw: true, element });
        } else if (element.name === "w:oMath" || element.name === "w:oMathPara") {
            // Math equations — preserve as raw Element
            currentChildren.push({ $type: "math", element } as unknown as FileChildJson);
        } else {
            // Preserve unknown body-level elements as raw (customXml, etc.)
            currentChildren.push({ $raw: true, element });
        }
    }

    // Final section (using body-level sectPr or default)
    sections.push({
        properties: finalSectPr ? parseSectionProperties(finalSectPr) : undefined,
        children: currentChildren,
    });

    return sections;
}

/** Resolve header/footer reference IDs to actual paragraph content */
function resolveHeadersFooters(sections: SectionJson[], ctx: DocxParseContext): void {
    for (const section of sections) {
        const props = section.properties as Record<string, unknown> | undefined;
        if (!props) continue;

        const headerRefs = props.headerRefs as Record<string, string> | undefined;
        const footerRefs = props.footerRefs as Record<string, string> | undefined;

        if (headerRefs) {
            section.headers = {};
            for (const [type, rId] of Object.entries(headerRefs)) {
                const content = parseHeaderFooterContent(rId, ctx);
                if (content) {
                    (section.headers as Record<string, { children: FileChildJson[] }>)[type] = {
                        children: content,
                    };
                }
            }
        }

        if (footerRefs) {
            section.footers = {};
            for (const [type, rId] of Object.entries(footerRefs)) {
                const content = parseHeaderFooterContent(rId, ctx);
                if (content) {
                    (section.footers as Record<string, { children: FileChildJson[] }>)[type] = {
                        children: content,
                    };
                }
            }
        }
    }
}

function parseHeaderFooterContent(rId: string, ctx: DocxParseContext): FileChildJson[] | undefined {
    const rel = ctx.documentRels.get(rId);
    if (!rel) return undefined;

    // Resolve path: rel.target is relative to word/, e.g. "header1.xml"
    const path = rel.target.startsWith("../")
        ? rel.target.replace("../", "word/")
        : `word/${rel.target}`;

    const xml = readXmlFromZip(ctx.zip, path);
    if (!xml) return undefined;

    const children: FileChildJson[] = [];
    for (const child of xml.elements ?? []) {
        if (child.name === "w:p") {
            children.push(parseParagraph(child, ctx));
        } else if (child.name === "w:tbl") {
            children.push(parseTable(child, ctx));
        }
    }
    return children.length > 0 ? children : undefined;
}

export type {
    DocxDocumentJson,
    SectionJson,
    FileChildJson,
    ParagraphJson,
    TextRunJson,
    ImageRunJson,
    ExternalHyperlinkJson,
    PageBreakJson,
    LineBreakJson,
    ColumnBreakJson,
    TableJson,
    TableRowJson,
    TableCellJson,
    SdtJson,
    SdtRunJson,
    MathJson,
    FieldJson,
    FootnoteEntry,
    CommentEntry,
    RawElement,
    MixedChildren,
} from "./types";
export { isRaw } from "./types";
