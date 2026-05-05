import * as fs from "fs";
import * as path from "path";

import {
    Document,
    Paragraph,
    TextRun,
    Run,
    Table,
    TableRow,
    TableCell,
    Packer,
    parseDocx,
    toDocumentOptions,
    HeadingLevel,
    AlignmentType,
    ExternalHyperlink,
    Header,
    Footer,
    ImageRun,
    BorderStyle,
    PageBreak,
    Tab,
    TabStopPosition,
    TabStopType,
    SimpleField,
    ShadingType,
    convertInchesToTwip,
    LeaderType,
} from "@office-open/docx";
import type { ParagraphJson, TextRunJson, TableJson, ImageRunJson } from "@office-open/docx";

function asParagraph(v: unknown): ParagraphJson {
    return v as ParagraphJson;
}

function asRun(v: unknown): TextRunJson {
    return v as TextRunJson;
}

function asTable(v: unknown): TableJson {
    return v as TableJson;
}

// Load images from demo/images/
const DEMO_IMAGES = path.join(import.meta.dirname, "images");
const CAT_JPG = fs.readFileSync(path.join(DEMO_IMAGES, "cat.jpg"));

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
                headers: {
                    default: new Header({
                        children: [
                            new Paragraph({
                                alignment: AlignmentType.RIGHT,
                                children: [
                                    new TextRun({
                                        text: "Header - Round-trip Test",
                                        size: 18,
                                        color: "888888",
                                    }),
                                ],
                            }),
                        ],
                    }),
                },
                footers: {
                    default: new Footer({
                        children: [
                            new Paragraph({
                                alignment: AlignmentType.CENTER,
                                children: [new TextRun({ text: "Page Footer" })],
                            }),
                        ],
                    }),
                },
                children: [
                    // ── Headings ──
                    new Paragraph({
                        heading: HeadingLevel.HEADING_1,
                        children: [new TextRun({ text: "Heading 1", bold: true, size: 56 })],
                    }),
                    new Paragraph({
                        heading: HeadingLevel.HEADING_2,
                        children: [new TextRun({ text: "Heading 2", bold: true, size: 36 })],
                    }),
                    new Paragraph({
                        heading: HeadingLevel.HEADING_3,
                        children: [new TextRun({ text: "Heading 3", bold: true, size: 28 })],
                    }),

                    // ── Rich text paragraph ──
                    new Paragraph({
                        spacing: { after: 200 },
                        children: [
                            new TextRun({ text: "This is a " }),
                            new TextRun({ text: "rich text", italics: true, color: "FF0000" }),
                            new TextRun({ text: " paragraph." }),
                        ],
                    }),

                    // ── Plain text paragraph ──
                    new Paragraph({
                        text: "Simple paragraph with plain text.",
                    }),

                    // ── Advanced text formatting ──
                    new Paragraph({
                        spacing: { after: 100 },
                        children: [
                            new TextRun({ text: "Bold", bold: true }),
                            new TextRun({ text: ", Italic", italics: true }),
                            new TextRun({ text: ", Underline", underline: { type: "single" } }),
                            new TextRun({ text: ", Strike", strike: true }),
                            new TextRun({ text: ", DoubleStrike", doubleStrike: true }),
                        ],
                    }),
                    new Paragraph({
                        spacing: { after: 100 },
                        children: [
                            new TextRun({ text: "SmallCaps", smallCaps: true }),
                            new TextRun({ text: " | AllCaps", allCaps: true }),
                            new TextRun({ text: " | Sub", subScript: true }),
                            new TextRun({ text: " | Super", superScript: true }),
                        ],
                    }),
                    new Paragraph({
                        spacing: { after: 100 },
                        children: [
                            new TextRun({ text: "Highlight", highlight: "yellow" }),
                            new TextRun({ text: " | Spacing", characterSpacing: 120 }),
                            new TextRun({ text: " | Shadow", shadow: true }),
                            new TextRun({ text: " | Outline", outline: true }),
                            new TextRun({ text: " | Emboss", emboss: true }),
                        ],
                    }),

                    // ── Alignment ──
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [new TextRun({ text: "Centered text" })],
                    }),
                    new Paragraph({
                        alignment: AlignmentType.RIGHT,
                        children: [new TextRun({ text: "Right-aligned text" })],
                    }),
                    new Paragraph({
                        alignment: AlignmentType.JUSTIFIED,
                        children: [
                            new TextRun({
                                text: "Justified text that should be long enough to show justification effect in the document layout.",
                            }),
                        ],
                    }),

                    // ── Hyperlink ──
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
                            new TextRun({ text: " for more info." }),
                        ],
                    }),

                    // ── Paragraph properties ──
                    new Paragraph({
                        indent: { left: 720, firstLine: 480 },
                        keepNext: true,
                        keepLines: true,
                        children: [
                            new TextRun({
                                text: "Indented paragraph with keepNext and keepLines.",
                            }),
                        ],
                    }),
                    new Paragraph({
                        indent: { left: 720, hanging: 360 },
                        children: [new TextRun({ text: "Hanging indent paragraph." })],
                    }),

                    // ── Tab stops ──
                    new Paragraph({
                        tabStops: [
                            {
                                position: TabStopPosition.MAX,
                                type: TabStopType.RIGHT,
                                leader: LeaderType.DOT,
                            },
                        ],
                        children: [
                            new TextRun({ text: "Tab leader" }),
                            new Run({ children: [new Tab()] }),
                            new TextRun({ text: "End" }),
                        ],
                    }),

                    // ── Page break ──
                    new Paragraph({
                        pageBreakBefore: true,
                        children: [new TextRun({ text: "Page 2 content" })],
                    }),
                    new Paragraph({
                        children: [new PageBreak()],
                    }),

                    // ── Column break ──
                    new Paragraph({
                        children: [new TextRun({ text: "After column break" })],
                    }),

                    // ── Simple field ──
                    new Paragraph({
                        children: [new SimpleField(' DATE \\@ "yyyy-MM-dd" ', "2026-01-01")],
                    }),

                    // ── Shading ──
                    new Paragraph({
                        shading: { type: ShadingType.SOLID, fill: "FFF2CC" },
                        children: [new TextRun({ text: "Shaded paragraph background" })],
                    }),

                    // ── Image ──
                    new Paragraph({
                        children: [
                            new TextRun({ text: "Image: " }),
                            new ImageRun({
                                data: CAT_JPG,
                                transformation: { width: 200, height: 150 },
                                type: "jpg",
                            }),
                        ],
                    }),

                    // ── Table with borders ──
                    new Table({
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [new Paragraph({ text: "Cell A1" })],
                                        width: { size: 3000, type: "dxa" },
                                        borders: {
                                            top: {
                                                style: BorderStyle.SINGLE,
                                                size: 6,
                                                color: "000000",
                                            },
                                            bottom: {
                                                style: BorderStyle.SINGLE,
                                                size: 6,
                                                color: "000000",
                                            },
                                            left: {
                                                style: BorderStyle.SINGLE,
                                                size: 6,
                                                color: "000000",
                                            },
                                            right: {
                                                style: BorderStyle.SINGLE,
                                                size: 6,
                                                color: "000000",
                                            },
                                        },
                                    }),
                                    new TableCell({
                                        children: [new Paragraph({ text: "Cell B1" })],
                                        width: { size: 3000, type: "dxa" },
                                        shading: { type: ShadingType.SOLID, fill: "D9E2F3" },
                                    }),
                                ],
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [new Paragraph({ text: "Cell A2" })],
                                        shading: { type: ShadingType.SOLID, fill: "E2EFDA" },
                                    }),
                                    new TableCell({
                                        children: [new Paragraph({ text: "Cell B2" })],
                                    }),
                                ],
                            }),
                        ],
                        width: { size: 6000, type: "dxa" },
                        borders: {
                            top: { style: BorderStyle.DOUBLE, size: 8, color: "2F5496" },
                            bottom: { style: BorderStyle.DOUBLE, size: 8, color: "2F5496" },
                            left: { style: BorderStyle.DOUBLE, size: 8, color: "2F5496" },
                            right: { style: BorderStyle.DOUBLE, size: 8, color: "2F5496" },
                            insideHorizontal: {
                                style: BorderStyle.SINGLE,
                                size: 4,
                                color: "2F5496",
                            },
                            insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "2F5496" },
                        },
                    }),

                    // ── Table with merged cells ──
                    new Table({
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [new Paragraph({ text: "Merged (2x1)" })],
                                        columnSpan: 2,
                                    }),
                                ],
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [new Paragraph({ text: "A2" })],
                                    }),
                                    new TableCell({
                                        children: [new Paragraph({ text: "B2" })],
                                    }),
                                ],
                            }),
                        ],
                    }),

                    // ── Numbered list ──
                    new Paragraph({
                        numbering: { reference: "my-numbering", level: 0 },
                        children: [new TextRun({ text: "Numbered item 1" })],
                    }),
                    new Paragraph({
                        numbering: { reference: "my-numbering", level: 0 },
                        children: [new TextRun({ text: "Numbered item 2" })],
                    }),
                    new Paragraph({
                        numbering: { reference: "my-numbering", level: 1 },
                        children: [new TextRun({ text: "Sub-item 2.1" })],
                    }),

                    // ── Bullet list ──
                    new Paragraph({
                        bullet: { level: 0 },
                        children: [new TextRun({ text: "Bullet item 1" })],
                    }),
                    new Paragraph({
                        bullet: { level: 0 },
                        children: [new TextRun({ text: "Bullet item 2" })],
                    }),
                    new Paragraph({
                        bullet: { level: 1 },
                        children: [new TextRun({ text: "Sub-bullet" })],
                    }),
                ],
            },
        ],
        numbering: {
            config: [
                {
                    reference: "my-numbering",
                    levels: [
                        {
                            level: 0,
                            format: "decimal",
                            text: "%1.",
                            alignment: AlignmentType.START,
                            style: {
                                paragraph: {
                                    indent: {
                                        left: convertInchesToTwip(0.5),
                                        hanging: convertInchesToTwip(0.25),
                                    },
                                },
                            },
                        },
                        {
                            level: 1,
                            format: "lowerLetter",
                            text: "%2)",
                            alignment: AlignmentType.START,
                            style: {
                                paragraph: {
                                    indent: {
                                        left: convertInchesToTwip(1),
                                        hanging: convertInchesToTwip(0.25),
                                    },
                                },
                            },
                        },
                    ],
                },
            ],
        },
    });

    // 2. Export to buffer
    const buffer = await Packer.toBuffer(doc);
    console.log(`Generated DOCX: ${buffer.length} bytes`);

    // 3. Parse it back
    const json = await parseDocx(new Uint8Array(buffer));

    // 4. Verify parse results
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

    console.log("\n--- Phase 1: Parse Verification ---");

    // Core properties
    console.log("\n--- Core Properties ---");
    assert("title", json.title === "Round-trip Test");
    assert("creator", json.creator === "Parser Demo");

    // Sections
    console.log("\n--- Sections ---");
    assert("1 section", json.sections.length === 1);
    const sec = json.sections[0];
    const page = (sec.properties as Record<string, unknown>)?.page as Record<string, unknown>;
    const pageSize = page?.size as Record<string, unknown>;
    const pageMargin = page?.margin as Record<string, unknown>;
    assert("page width", pageSize?.width === 11906);
    assert("page height", pageSize?.height === 16838);
    assert("margin top", pageMargin?.top === 1440);

    // Count children (header/footer not in children)
    const children = sec.children;
    console.log(`\n--- Paragraphs & Elements (${children.length} children) ---`);

    // Para 1: Heading 1
    const p1 = asParagraph(children[0]);
    assert("para 0 heading style", p1.style === "Heading1");
    assert("para 0 run bold", asRun(p1.children?.[0]).bold === true);

    // Para 2: Heading 2
    const p2 = asParagraph(children[1]);
    assert("para 1 heading style", p2.style === "Heading2");

    // Para 3: Heading 3
    const p3 = asParagraph(children[2]);
    assert("para 2 heading style", p3.style === "Heading3");

    // Para 4: Rich text
    const p4 = asParagraph(children[3]);
    assert("para 3 spacing.after", p4.spacing?.after === 200);
    assert("para 3 has 3 runs", p4.children?.length === 3);
    assert("para 3 run 2 italics", asRun(p4.children?.[1]).italics === true);
    assert("para 3 run 2 color", asRun(p4.children?.[1]).color === "FF0000");

    // Para 5: Plain text
    const p5 = asParagraph(children[4]);
    assert(
        "para 4 plain text",
        asRun(p5.children?.[0]).text === "Simple paragraph with plain text.",
    );

    // Para 6: Multiple formatting
    const p6 = asParagraph(children[5]);
    assert("para 5 has 5 runs", p6.children?.length === 5);
    assert("para 5 run 1 bold", asRun(p6.children?.[0]).bold === true);
    assert("para 5 run 2 italics", asRun(p6.children?.[1]).italics === true);
    assert("para 5 run 3 underline type", asRun(p6.children?.[2]).underline?.type === "single");
    assert("para 5 run 4 strike", asRun(p6.children?.[3]).strike === true);
    assert("para 5 run 5 doubleStrike", asRun(p6.children?.[4]).doubleStrike === true);

    // Para 7: Advanced formatting (smallCaps, allCaps, sub, super)
    const p7 = asParagraph(children[6]);
    assert("para 6 run 1 smallCaps", asRun(p7.children?.[0]).smallCaps === true);
    assert("para 6 run 2 allCaps", asRun(p7.children?.[1]).allCaps === true);
    assert("para 6 run 3 subScript", asRun(p7.children?.[2]).subScript === true);
    assert("para 6 run 4 superScript", asRun(p7.children?.[3]).superScript === true);

    // Para 8: highlight, spacing, shadow, outline, emboss
    const p8 = asParagraph(children[7]);
    assert("para 7 run 1 highlight", asRun(p8.children?.[0]).highlight === "yellow");
    assert("para 7 run 2 characterSpacing", asRun(p8.children?.[1]).characterSpacing === 120);
    assert("para 7 run 3 shadow", asRun(p8.children?.[2]).shadow === true);
    assert("para 7 run 4 outline", asRun(p8.children?.[3]).outline === true);
    assert("para 7 run 5 emboss", asRun(p8.children?.[4]).emboss === true);

    // Alignment
    const p9 = asParagraph(children[8]);
    assert("para 8 center alignment", p9.alignment === "center");
    const p10 = asParagraph(children[9]);
    assert("para 9 right alignment", p10.alignment === "right");

    // Hyperlink
    const p11 = asParagraph(children[11]);
    const hlink = p11.children?.find(
        (c) => (c as Record<string, unknown>).$type === "externalHyperlink",
    );
    assert("para 11 has hyperlink", !!hlink);
    assert("hyperlink url", (hlink as Record<string, unknown>)?.link === "https://example.com");

    // Indent
    const p12 = asParagraph(children[12]);
    assert("para 12 indent.left", p12.indent?.left === 720);
    assert("para 12 indent.firstLine", p12.indent?.firstLine === 480);
    assert("para 12 keepNext", p12.keepNext === true);
    assert("para 12 keepLines", p12.keepLines === true);

    // Table with borders
    const tblIdx = children.findIndex((c) => (c as Record<string, unknown>).$type === "table");
    assert("has table", tblIdx >= 0);
    const tbl = asTable(children[tblIdx]);
    assert("table has 2 rows", tbl.rows.length === 2);
    assert("row 1 has 2 cells", tbl.rows[0].cells.length === 2);
    assert(
        "cell A1 text",
        asRun((tbl.rows[0].cells[0].children[0] as ParagraphJson).children?.[0]).text === "Cell A1",
    );

    // Verify table borders are in IBorderOptions format (not raw XML attributes)
    const tblBorders = tbl.borders as Record<string, unknown> | undefined;
    assert("table borders exist", !!tblBorders);
    if (tblBorders) {
        const topBorder = tblBorders.top as Record<string, unknown>;
        assert("table top border has style (not w:val)", !!topBorder?.style);
        assert("table top border no w:val key", !("w:val" in (topBorder ?? {})));
        assert("table top border style is double", topBorder?.style === "double");
        assert("table top border size is 8", topBorder?.size === 8);
        assert("table top border color is 2F5496", topBorder?.color === "2F5496");
    }

    // Verify cell borders are in IBorderOptions format
    const cellBorders = tbl.rows[0].cells[0].borders as Record<string, unknown> | undefined;
    assert("cell A1 borders exist", !!cellBorders);
    if (cellBorders) {
        const cellTop = cellBorders.top as Record<string, unknown>;
        assert("cell A1 top border has style (not w:val)", !!cellTop?.style);
        assert("cell A1 top border no w:val key", !("w:val" in (cellTop ?? {})));
    }

    // Merged cells table
    const tblIdx2 = children.findIndex(
        (c, i) => i > tblIdx && (c as Record<string, unknown>).$type === "table",
    );
    assert("has second table", tblIdx2 >= 0);
    const tbl2 = asTable(children[tblIdx2]);
    assert("table 2 row 1 columnSpan", tbl2.rows[0].cells[0].columnSpan === 2);

    // ImageRun (parser stores image data; verify existence)
    // Search within paragraph children, not section children
    let foundImage = false;
    for (const child of children) {
        if ((child as Record<string, unknown>).$type === "paragraph") {
            const p = child as ParagraphJson;
            for (const pc of p.children ?? []) {
                if ((pc as Record<string, unknown>).$type === "imageRun") {
                    foundImage = true;
                    const img = pc as ImageRunJson;
                    assert("image type is jpg", img.type === "jpg");
                    assert("image transformation exists", !!img.transformation);
                    break;
                }
            }
            if (foundImage) break;
        }
    }
    if (!foundImage) {
        console.log("  SKIP: imageRun not found in paragraph children");
    }

    // Numbering (parser extracts numbering.xml and builds config)
    console.log("\n--- Numbering ---");
    assert("numbering config exists", !!json.numbering);
    if (json.numbering) {
        assert("numbering config has entries", json.numbering.length >= 1);
        console.log(`  numbering config entries: ${json.numbering.length}`);
    }

    // Headers/Footers
    console.log("\n--- Headers/Footers ---");
    const secHeaders = sec.headers;
    const secFooters = sec.footers;
    assert("section has headers", !!secHeaders && Object.keys(secHeaders).length > 0);
    assert("section has footers", !!secFooters && Object.keys(secFooters).length > 0);
    console.log(`  header types: ${Object.keys(secHeaders as Record<string, unknown>).join(", ")}`);
    console.log(`  footer types: ${Object.keys(secFooters as Record<string, unknown>).join(", ")}`);

    // 5. Round-trip: convert parsed JSON → constructor instances → Packer.toBuffer()
    console.log("\n--- Phase 2: Convert & Re-pack ---");
    const roundTripDoc = new Document(toDocumentOptions(json));
    const roundTripBuffer = await Packer.toBuffer(roundTripDoc);
    assert("round-trip buffer non-empty", roundTripBuffer.length > 0);
    console.log(`Round-trip DOCX: ${roundTripBuffer.length} bytes`);

    // 6. Re-parse round-trip buffer
    const json2 = await parseDocx(new Uint8Array(roundTripBuffer));

    // 7. Verify re-parsed data matches
    console.log("\n--- Phase 3: Re-parse Verification ---");
    assert("re-parsed title", json2.title === "Round-trip Test");
    assert("re-parsed creator", json2.creator === "Parser Demo");
    assert("re-parsed 1 section", json2.sections.length === 1);
    assert("re-parsed children count", json2.sections[0].children.length === children.length);

    const sec2 = json2.sections[0];
    const rp1 = asParagraph(sec2.children[0]);
    assert("re-parsed heading style", rp1.style === "Heading1");
    assert("re-parsed heading bold", asRun(rp1.children?.[0]).bold === true);

    const rp5 = asParagraph(sec2.children[4]);
    assert(
        "re-parsed plain text",
        asRun(rp5.children?.[0]).text === "Simple paragraph with plain text.",
    );

    // Re-parsed table borders should still be IBorderOptions format
    const rpTblIdx = sec2.children.findIndex(
        (c) => (c as Record<string, unknown>).$type === "table",
    );
    const rpTbl = asTable(sec2.children[rpTblIdx]);
    const rpTblBorders = rpTbl.borders as Record<string, unknown> | undefined;
    if (rpTblBorders) {
        const rpTopBorder = rpTblBorders.top as Record<string, unknown>;
        assert("re-parsed table top border style", rpTopBorder?.style === "double");
        assert("re-parsed table top border size", rpTopBorder?.size === 8);
        assert("re-parsed table top border color", rpTopBorder?.color === "2F5496");
    }

    // Summary
    console.log(`\n=== Results: ${pass} passed, ${fail} failed ===`);
    if (fail > 0) process.exit(1);

    // Save to files
    fs.writeFileSync("My Document.docx", buffer);
    fs.writeFileSync("My Document (round-trip).docx", roundTripBuffer);
    console.log("\nSaved original and round-trip DOCX files");
}

main().catch(console.error);
