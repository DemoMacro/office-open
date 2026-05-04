import * as fs from "fs";

import {
    Document,
    Paragraph,
    TextRun,
    Table,
    TableRow,
    TableCell,
    Packer,
    parseDocx,
    HeadingLevel,
    AlignmentType,
    ExternalHyperlink,
} from "@office-open/docx";
import type { ParagraphJson, TextRunJson, TableJson } from "@office-open/docx";

function asParagraph(v: unknown): ParagraphJson {
    return v as ParagraphJson;
}

function asRun(v: unknown): TextRunJson {
    return v as TextRunJson;
}

function asTable(v: unknown): TableJson {
    return v as TableJson;
}

async function main() {
    // 1. Create a document with diverse content
    const doc = new Document({
        title: "Round-trip Test",
        creator: "Parser Demo",
        subject: "Parser Verification",
        sections: [
            {
                properties: {
                    page: {
                        size: { width: 11906, height: 16838 },
                        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
                    },
                },
                children: [
                    // Heading
                    new Paragraph({
                        heading: HeadingLevel.HEADING_1,
                        children: [new TextRun({ text: "Hello World", bold: true, size: 56 })],
                    }),
                    // Rich text paragraph
                    new Paragraph({
                        spacing: { after: 200 },
                        children: [
                            new TextRun({ text: "This is a " }),
                            new TextRun({ text: "rich text", italics: true, color: "FF0000" }),
                            new TextRun({ text: " paragraph." }),
                        ],
                    }),
                    // Plain text paragraph
                    new Paragraph({
                        text: "Simple paragraph with plain text.",
                    }),
                    // Multiple formatting
                    new Paragraph({
                        children: [
                            new TextRun({ text: "Bold", bold: true }),
                            new TextRun({ text: ", Italic", italics: true }),
                            new TextRun({ text: ", Underline", underline: { type: "single" } }),
                            new TextRun({ text: ", Strike", strike: true }),
                        ],
                    }),
                    // Hyperlink
                    new Paragraph({
                        children: [
                            new TextRun({ text: "Visit " }),
                            new ExternalHyperlink({
                                link: "https://example.com",
                                children: [
                                    new TextRun({
                                        text: "example.com",
                                        color: "0563C1",
                                        underline: { type: "single" },
                                    }),
                                ],
                            }),
                        ],
                    }),
                    // Alignment
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [new TextRun({ text: "Centered text" })],
                    }),
                    // Simple table
                    new Table({
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [new Paragraph({ text: "Cell A1" })],
                                        width: { size: 3000, type: "dxa" },
                                    }),
                                    new TableCell({
                                        children: [new Paragraph({ text: "Cell B1" })],
                                        width: { size: 3000, type: "dxa" },
                                    }),
                                ],
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [new Paragraph({ text: "Cell A2" })],
                                    }),
                                    new TableCell({
                                        children: [new Paragraph({ text: "Cell B2" })],
                                    }),
                                ],
                            }),
                        ],
                    }),
                ],
            },
        ],
    });

    // 2. Export to buffer
    const buffer = await Packer.toBuffer(doc);
    console.log(`Generated DOCX: ${buffer.length} bytes`);

    // 3. Parse it back
    const json = await parseDocx(new Uint8Array(buffer));

    // 4. Verify round-trip
    let pass = 0;
    let fail = 0;
    const assert = (label: string, condition: boolean) => {
        if (condition) {
            pass++;
            console.log(`  PASS: ${label}`);
        } else {
            fail++;
            console.log(`  FAIL: ${label}`);
        }
    };

    console.log("\n--- Core Properties ---");
    assert("title", json.title === "Round-trip Test");
    assert("creator", json.creator === "Parser Demo");

    console.log("\n--- Sections ---");
    assert("1 section", json.sections.length === 1);

    const sec = json.sections[0];
    const page = (sec.properties as Record<string, unknown>)?.page as Record<string, unknown>;
    const pageSize = page?.size as Record<string, unknown>;
    const pageMargin = page?.margin as Record<string, unknown>;
    assert("page width", pageSize?.width === 11906);
    assert("page height", pageSize?.height === 16838);
    assert("margin top", pageMargin?.top === 1440);

    console.log("\n--- Paragraphs ---");
    assert("7 children", sec.children.length === 7);

    // Para 1: Heading
    const p1 = asParagraph(sec.children[0]);
    assert("para 1 is paragraph", p1.$type === "paragraph");
    assert("para 1 heading style", p1.style === "Heading1");
    const p1Run = asRun(p1.children?.[0]);
    assert("para 1 run text", p1Run.text === "Hello World");
    assert("para 1 run bold", p1Run.bold === true);
    assert("para 1 run size", p1Run.size === 56);

    // Para 2: Rich text
    const p2 = asParagraph(sec.children[1]);
    assert("para 2 spacing.after", p2.spacing?.after === 200);
    assert("para 2 has 3 runs", p2.children?.length === 3);
    const p2Run2 = asRun(p2.children?.[1]);
    assert("para 2 run 2 italics", p2Run2.italics === true);
    assert("para 2 run 2 color", p2Run2.color === "FF0000");

    // Para 3: Plain text
    const p3 = asParagraph(sec.children[2]);
    assert("para 3 plain text", p3.text === "Simple paragraph with plain text.");

    // Para 4: Multiple formatting
    const p4 = asParagraph(sec.children[3]);
    assert("para 4 has 4 runs", p4.children?.length === 4);
    assert("para 4 run 1 bold", asRun(p4.children?.[0]).bold === true);
    assert("para 4 run 2 italics", asRun(p4.children?.[1]).italics === true);
    assert("para 4 run 3 underline type", asRun(p4.children?.[2]).underline?.type === "single");
    assert("para 4 run 4 strike", asRun(p4.children?.[3]).strike === true);

    // Para 5: Hyperlink
    const p5 = asParagraph(sec.children[4]);
    assert(
        "para 5 has hyperlink",
        p5.children?.some((c) => c.$type === "externalHyperlink") === true,
    );
    const hlink = p5.children?.find((c) => c.$type === "externalHyperlink");
    assert("hyperlink url", (hlink as Record<string, unknown>)?.link === "https://example.com");

    // Para 6: Alignment
    const p6 = asParagraph(sec.children[5]);
    assert("para 6 alignment", p6.alignment === "center");

    // Para 7: Table
    const tbl = asTable(sec.children[6]);
    assert("para 7 is table", tbl.$type === "table");
    assert("table has 2 rows", tbl.rows.length === 2);
    assert("row 1 has 2 cells", tbl.rows[0].cells.length === 2);
    assert("cell A1 text", tbl.rows[0].cells[0].children[0].text === "Cell A1");

    // Summary
    console.log(`\n=== Results: ${pass} passed, ${fail} failed ===`);
    if (fail > 0) process.exit(1);

    // Save to file
    fs.writeFileSync("My Document.docx", buffer);
    console.log("\nSaved to My Document.docx");
}

main().catch(console.error);
