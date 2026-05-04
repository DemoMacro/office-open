import { unzipToMap, readXmlFromZip } from "@office-open/core";
import { parseCoreProperties } from "@office-open/core";
import { findChild, type Element } from "@office-open/xml";

import { createDocxParseContext } from "./context";
import { parseParagraph } from "./paragraph";
import { parseSectionProperties } from "./section";
import { parseTable } from "./table";
import type { DocxDocumentJson, SectionJson, FileChildJson } from "./types";

export async function parseDocx(data: Uint8Array): Promise<DocxDocumentJson> {
    const zip = unzipToMap(data);
    const ctx = createDocxParseContext(zip);

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

    return {
        sections,
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
        }
        // Skip other body-level elements (sdt, customXml, etc.) for MVP
    }

    // Final section (using body-level sectPr or default)
    sections.push({
        properties: finalSectPr ? parseSectionProperties(finalSectPr) : undefined,
        children: currentChildren,
    });

    return sections;
}

export type {
    DocxDocumentJson,
    SectionJson,
    FileChildJson,
    ParagraphJson,
    TextRunJson,
    ImageRunJson,
    ExternalHyperlinkJson,
    TableJson,
    TableRowJson,
    TableCellJson,
} from "./types";
