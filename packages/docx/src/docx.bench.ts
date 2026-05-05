import {
    AlignmentType as AlignmentTypeOrig,
    Document as DocumentOrig,
    Footer as FooterOrig,
    Header as HeaderOrig,
    HeadingLevel as HeadingLevelOrig,
    PageNumber as PageNumberOrig,
    Paragraph as ParagraphOrig,
    Packer as PackerOrig,
    Table as TableOrig,
    TableCell as TableCellOrig,
    TableRow as TableRowOrig,
    TextRun as TextRunOrig,
    UnderlineType as UnderlineTypeOrig,
} from "docx";
import { bench, describe } from "vite-plus/test";

import {
    AlignmentType,
    Document,
    Footer,
    Header,
    HeadingLevel,
    PageNumber,
    Packer,
    Paragraph,
    Table,
    TableCell,
    TableRow,
    TextRun,
    UnderlineType,
    WidthType,
} from "./index";

// ── Shared fixture data ──

const PARAGRAPH_CHILDREN = Array.from({ length: 20 }, (_, i) => ({
    text: `Paragraph text line ${i + 1} with some content`,
    bold: i % 3 === 0,
    italics: i % 5 === 0,
}));

const TABLE_ROWS = Array.from({ length: 10 }, (_, rowIdx) => ({
    cells: Array.from({ length: 5 }, (_, colIdx) => ({
        text: `R${rowIdx + 1}C${colIdx + 1}`,
        width: { size: 2000, type: WidthType.PERCENTAGE },
    })),
}));

// ── Benchmarks ──

describe("DOCX: Object Creation (no pack)", () => {
    bench("ours — simple", () => {
        new Document({
            sections: [
                {
                    children: [
                        new Paragraph({ children: [new TextRun("Hello World")] }),
                        new Paragraph({ children: [new TextRun("Second paragraph")] }),
                    ],
                },
            ],
        });
    });

    bench("original — simple", () => {
        new DocumentOrig({
            sections: [
                {
                    children: [
                        new ParagraphOrig({ children: [new TextRunOrig("Hello World")] }),
                        new ParagraphOrig({ children: [new TextRunOrig("Second paragraph")] }),
                    ],
                },
            ],
        });
    });

    bench("ours — styled paragraphs (20)", () => {
        new Document({
            sections: [
                {
                    children: PARAGRAPH_CHILDREN.map(
                        (p) =>
                            new Paragraph({
                                children: [
                                    new TextRun({ text: p.text, bold: p.bold, italics: p.italics }),
                                ],
                            }),
                    ),
                },
            ],
        });
    });

    bench("original — styled paragraphs (20)", () => {
        new DocumentOrig({
            sections: [
                {
                    children: PARAGRAPH_CHILDREN.map(
                        (p) =>
                            new ParagraphOrig({
                                children: [
                                    new TextRunOrig({
                                        text: p.text,
                                        bold: p.bold,
                                        italics: p.italics,
                                    }),
                                ],
                            }),
                    ),
                },
            ],
        });
    });

    bench("ours — table (10x5)", () => {
        new Document({
            sections: [
                {
                    children: [
                        new Table({
                            rows: TABLE_ROWS.map(
                                (row) =>
                                    new TableRow({
                                        children: row.cells.map(
                                            (cell) =>
                                                new TableCell({
                                                    width: {
                                                        size: cell.width.size,
                                                        type: cell.width.type,
                                                    },
                                                    children: [
                                                        new Paragraph({
                                                            children: [new TextRun(cell.text)],
                                                        }),
                                                    ],
                                                }),
                                        ),
                                    }),
                            ),
                        }),
                    ],
                },
            ],
        });
    });

    bench("original — table (10x5)", () => {
        new DocumentOrig({
            sections: [
                {
                    children: [
                        new TableOrig({
                            rows: TABLE_ROWS.map(
                                (row) =>
                                    new TableRowOrig({
                                        children: row.cells.map(
                                            (cell) =>
                                                new TableCellOrig({
                                                    width: {
                                                        size: cell.width.size,
                                                        type: cell.width.type,
                                                    },
                                                    children: [
                                                        new ParagraphOrig({
                                                            children: [new TextRunOrig(cell.text)],
                                                        }),
                                                    ],
                                                }),
                                        ),
                                    }),
                            ),
                        }),
                    ],
                },
            ],
        });
    });

    bench("ours — full featured", () => {
        new Document({
            sections: [
                {
                    headers: {
                        default: new Header({
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.RIGHT,
                                    children: [
                                        new TextRun("Header text"),
                                        new TextRun({ children: [PageNumber.CURRENT] }),
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
                                    children: [new TextRun("Footer")],
                                }),
                            ],
                        }),
                    },
                    children: [
                        new Paragraph({
                            heading: HeadingLevel.HEADING_1,
                            children: [new TextRun("Document Title")],
                        }),
                        new Paragraph({
                            heading: HeadingLevel.HEADING_2,
                            children: [new TextRun("Section 1")],
                        }),
                        ...PARAGRAPH_CHILDREN.map(
                            (p) =>
                                new Paragraph({
                                    children: [
                                        new TextRun({
                                            text: p.text,
                                            bold: p.bold,
                                            italics: p.italics,
                                        }),
                                    ],
                                }),
                        ),
                        new Table({
                            rows: TABLE_ROWS.map(
                                (row) =>
                                    new TableRow({
                                        children: row.cells.map(
                                            (cell) =>
                                                new TableCell({
                                                    width: {
                                                        size: cell.width.size,
                                                        type: cell.width.type,
                                                    },
                                                    children: [
                                                        new Paragraph({
                                                            children: [new TextRun(cell.text)],
                                                        }),
                                                    ],
                                                }),
                                        ),
                                    }),
                            ),
                        }),
                    ],
                },
            ],
        });
    });

    bench("original — full featured", () => {
        new DocumentOrig({
            sections: [
                {
                    headers: {
                        default: new HeaderOrig({
                            children: [
                                new ParagraphOrig({
                                    alignment: AlignmentTypeOrig.RIGHT,
                                    children: [
                                        new TextRunOrig("Header text"),
                                        new TextRunOrig({ children: [PageNumberOrig.CURRENT] }),
                                    ],
                                }),
                            ],
                        }),
                    },
                    footers: {
                        default: new FooterOrig({
                            children: [
                                new ParagraphOrig({
                                    alignment: AlignmentTypeOrig.CENTER,
                                    children: [new TextRunOrig("Footer")],
                                }),
                            ],
                        }),
                    },
                    children: [
                        new ParagraphOrig({
                            heading: HeadingLevelOrig.HEADING_1,
                            children: [new TextRunOrig("Document Title")],
                        }),
                        new ParagraphOrig({
                            heading: HeadingLevelOrig.HEADING_2,
                            children: [new TextRunOrig("Section 1")],
                        }),
                        ...PARAGRAPH_CHILDREN.map(
                            (p) =>
                                new ParagraphOrig({
                                    children: [
                                        new TextRunOrig({
                                            text: p.text,
                                            bold: p.bold,
                                            italics: p.italics,
                                        }),
                                    ],
                                }),
                        ),
                        new TableOrig({
                            rows: TABLE_ROWS.map(
                                (row) =>
                                    new TableRowOrig({
                                        children: row.cells.map(
                                            (cell) =>
                                                new TableCellOrig({
                                                    width: {
                                                        size: cell.width.size,
                                                        type: cell.width.type,
                                                    },
                                                    children: [
                                                        new ParagraphOrig({
                                                            children: [new TextRunOrig(cell.text)],
                                                        }),
                                                    ],
                                                }),
                                        ),
                                    }),
                            ),
                        }),
                    ],
                },
            ],
        });
    });
});

describe("DOCX: Create + toBuffer", () => {
    bench("ours — simple + toBuffer", async () => {
        const doc = new Document({
            sections: [
                {
                    children: [
                        new Paragraph({ children: [new TextRun("Hello World")] }),
                        new Paragraph({ children: [new TextRun("Second paragraph")] }),
                    ],
                },
            ],
        });
        await Packer.toBuffer(doc);
    });

    bench("original — simple + toBuffer", async () => {
        const doc = new DocumentOrig({
            sections: [
                {
                    children: [
                        new ParagraphOrig({ children: [new TextRunOrig("Hello World")] }),
                        new ParagraphOrig({ children: [new TextRunOrig("Second paragraph")] }),
                    ],
                },
            ],
        });
        await PackerOrig.toBuffer(doc);
    });

    bench("ours — styled paragraphs (20) + toBuffer", async () => {
        const doc = new Document({
            sections: [
                {
                    children: PARAGRAPH_CHILDREN.map(
                        (p) =>
                            new Paragraph({
                                children: [
                                    new TextRun({ text: p.text, bold: p.bold, italics: p.italics }),
                                ],
                            }),
                    ),
                },
            ],
        });
        await Packer.toBuffer(doc);
    });

    bench("original — styled paragraphs (20) + toBuffer", async () => {
        const doc = new DocumentOrig({
            sections: [
                {
                    children: PARAGRAPH_CHILDREN.map(
                        (p) =>
                            new ParagraphOrig({
                                children: [
                                    new TextRunOrig({
                                        text: p.text,
                                        bold: p.bold,
                                        italics: p.italics,
                                    }),
                                ],
                            }),
                    ),
                },
            ],
        });
        await PackerOrig.toBuffer(doc);
    });

    bench("ours — table (10x5) + toBuffer", async () => {
        const doc = new Document({
            sections: [
                {
                    children: [
                        new Table({
                            rows: TABLE_ROWS.map(
                                (row) =>
                                    new TableRow({
                                        children: row.cells.map(
                                            (cell) =>
                                                new TableCell({
                                                    width: {
                                                        size: cell.width.size,
                                                        type: cell.width.type,
                                                    },
                                                    children: [
                                                        new Paragraph({
                                                            children: [new TextRun(cell.text)],
                                                        }),
                                                    ],
                                                }),
                                        ),
                                    }),
                            ),
                        }),
                    ],
                },
            ],
        });
        await Packer.toBuffer(doc);
    });

    bench("original — table (10x5) + toBuffer", async () => {
        const doc = new DocumentOrig({
            sections: [
                {
                    children: [
                        new TableOrig({
                            rows: TABLE_ROWS.map(
                                (row) =>
                                    new TableRowOrig({
                                        children: row.cells.map(
                                            (cell) =>
                                                new TableCellOrig({
                                                    width: {
                                                        size: cell.width.size,
                                                        type: cell.width.type,
                                                    },
                                                    children: [
                                                        new ParagraphOrig({
                                                            children: [new TextRunOrig(cell.text)],
                                                        }),
                                                    ],
                                                }),
                                        ),
                                    }),
                            ),
                        }),
                    ],
                },
            ],
        });
        await PackerOrig.toBuffer(doc);
    });

    bench("ours — full featured + toBuffer", async () => {
        const doc = new Document({
            sections: [
                {
                    headers: {
                        default: new Header({
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.RIGHT,
                                    children: [
                                        new TextRun("Header text"),
                                        new TextRun({ children: [PageNumber.CURRENT] }),
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
                                    children: [new TextRun("Footer")],
                                }),
                            ],
                        }),
                    },
                    children: [
                        new Paragraph({
                            heading: HeadingLevel.HEADING_1,
                            children: [new TextRun("Document Title")],
                        }),
                        new Paragraph({
                            heading: HeadingLevel.HEADING_2,
                            children: [new TextRun("Section 1")],
                        }),
                        ...PARAGRAPH_CHILDREN.map(
                            (p) =>
                                new Paragraph({
                                    children: [
                                        new TextRun({
                                            text: p.text,
                                            bold: p.bold,
                                            italics: p.italics,
                                        }),
                                    ],
                                }),
                        ),
                        new Table({
                            rows: TABLE_ROWS.map(
                                (row) =>
                                    new TableRow({
                                        children: row.cells.map(
                                            (cell) =>
                                                new TableCell({
                                                    width: {
                                                        size: cell.width.size,
                                                        type: cell.width.type,
                                                    },
                                                    children: [
                                                        new Paragraph({
                                                            children: [new TextRun(cell.text)],
                                                        }),
                                                    ],
                                                }),
                                        ),
                                    }),
                            ),
                        }),
                    ],
                },
            ],
        });
        await Packer.toBuffer(doc);
    });

    bench("original — full featured + toBuffer", async () => {
        const doc = new DocumentOrig({
            sections: [
                {
                    headers: {
                        default: new HeaderOrig({
                            children: [
                                new ParagraphOrig({
                                    alignment: AlignmentTypeOrig.RIGHT,
                                    children: [
                                        new TextRunOrig("Header text"),
                                        new TextRunOrig({ children: [PageNumberOrig.CURRENT] }),
                                    ],
                                }),
                            ],
                        }),
                    },
                    footers: {
                        default: new FooterOrig({
                            children: [
                                new ParagraphOrig({
                                    alignment: AlignmentTypeOrig.CENTER,
                                    children: [new TextRunOrig("Footer")],
                                }),
                            ],
                        }),
                    },
                    children: [
                        new ParagraphOrig({
                            heading: HeadingLevelOrig.HEADING_1,
                            children: [new TextRunOrig("Document Title")],
                        }),
                        new ParagraphOrig({
                            heading: HeadingLevelOrig.HEADING_2,
                            children: [new TextRunOrig("Section 1")],
                        }),
                        ...PARAGRAPH_CHILDREN.map(
                            (p) =>
                                new ParagraphOrig({
                                    children: [
                                        new TextRunOrig({
                                            text: p.text,
                                            bold: p.bold,
                                            italics: p.italics,
                                        }),
                                    ],
                                }),
                        ),
                        new TableOrig({
                            rows: TABLE_ROWS.map(
                                (row) =>
                                    new TableRowOrig({
                                        children: row.cells.map(
                                            (cell) =>
                                                new TableCellOrig({
                                                    width: {
                                                        size: cell.width.size,
                                                        type: cell.width.type,
                                                    },
                                                    children: [
                                                        new ParagraphOrig({
                                                            children: [new TextRunOrig(cell.text)],
                                                        }),
                                                    ],
                                                }),
                                        ),
                                    }),
                            ),
                        }),
                    ],
                },
            ],
        });
        await PackerOrig.toBuffer(doc);
    });
});

// ── Large file benchmarks ──

const LARGE_PARAGRAPHS = Array.from({ length: 500 }, (_, i) => ({
    text: `This is paragraph ${i + 1} with sample content to simulate real document generation. Each paragraph contains enough text to represent a realistic scenario.`,
    bold: i % 3 === 0,
    italics: i % 5 === 0,
    underline: i % 7 === 0,
}));

const LARGE_TABLE_ROWS = Array.from({ length: 100 }, (_, rowIdx) => ({
    cells: Array.from({ length: 10 }, (_, colIdx) => ({
        text: `R${rowIdx + 1}C${colIdx + 1} data`,
        width: { size: 1000, type: WidthType.PERCENTAGE },
    })),
}));

describe("DOCX: Large Files — Create + toBuffer", () => {
    bench("ours — 500 paragraphs + toBuffer", async () => {
        const doc = new Document({
            sections: [
                {
                    children: LARGE_PARAGRAPHS.map(
                        (p) =>
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: p.text,
                                        bold: p.bold,
                                        italics: p.italics,
                                        underline: { type: UnderlineType.SINGLE },
                                    }),
                                ],
                            }),
                    ),
                },
            ],
        });
        await Packer.toBuffer(doc);
    });

    bench("original — 500 paragraphs + toBuffer", async () => {
        const doc = new DocumentOrig({
            sections: [
                {
                    children: LARGE_PARAGRAPHS.map(
                        (p) =>
                            new ParagraphOrig({
                                children: [
                                    new TextRunOrig({
                                        text: p.text,
                                        bold: p.bold,
                                        italics: p.italics,
                                        underline: { type: UnderlineTypeOrig.SINGLE },
                                    }),
                                ],
                            }),
                    ),
                },
            ],
        });
        await PackerOrig.toBuffer(doc);
    });

    bench("ours — 100x10 table + toBuffer", async () => {
        const doc = new Document({
            sections: [
                {
                    children: [
                        new Table({
                            rows: LARGE_TABLE_ROWS.map(
                                (row) =>
                                    new TableRow({
                                        children: row.cells.map(
                                            (cell) =>
                                                new TableCell({
                                                    width: {
                                                        size: cell.width.size,
                                                        type: cell.width.type,
                                                    },
                                                    children: [
                                                        new Paragraph({
                                                            children: [new TextRun(cell.text)],
                                                        }),
                                                    ],
                                                }),
                                        ),
                                    }),
                            ),
                        }),
                    ],
                },
            ],
        });
        await Packer.toBuffer(doc);
    });

    bench("original — 100x10 table + toBuffer", async () => {
        const doc = new DocumentOrig({
            sections: [
                {
                    children: [
                        new TableOrig({
                            rows: LARGE_TABLE_ROWS.map(
                                (row) =>
                                    new TableRowOrig({
                                        children: row.cells.map(
                                            (cell) =>
                                                new TableCellOrig({
                                                    width: {
                                                        size: cell.width.size,
                                                        type: cell.width.type,
                                                    },
                                                    children: [
                                                        new ParagraphOrig({
                                                            children: [new TextRunOrig(cell.text)],
                                                        }),
                                                    ],
                                                }),
                                        ),
                                    }),
                            ),
                        }),
                    ],
                },
            ],
        });
        await PackerOrig.toBuffer(doc);
    });

    bench("ours — 10 sections × 50 paragraphs + toBuffer", async () => {
        const doc = new Document({
            sections: Array.from({ length: 10 }, (_, si) => ({
                properties: { page: { margin: { top: 1440, bottom: 1440 } } },
                children: Array.from({ length: 50 }, (_, pi) => {
                    return new Paragraph({
                        heading: pi === 0 ? HeadingLevel.HEADING_1 : undefined,
                        children: [
                            new TextRun({
                                text:
                                    pi === 0
                                        ? `Chapter ${si + 1} Title`
                                        : `Chapter ${si + 1} paragraph ${pi} body content for realistic document simulation.`,
                                bold: pi === 0,
                            }),
                        ],
                    });
                }),
            })),
        });
        await Packer.toBuffer(doc);
    });

    bench("original — 10 sections × 50 paragraphs + toBuffer", async () => {
        const doc = new DocumentOrig({
            sections: Array.from({ length: 10 }, (_, si) => ({
                properties: { page: { margin: { top: 1440, bottom: 1440 } } },
                children: Array.from({ length: 50 }, (_, pi) => {
                    return new ParagraphOrig({
                        heading: pi === 0 ? HeadingLevelOrig.HEADING_1 : undefined,
                        children: [
                            new TextRunOrig({
                                text:
                                    pi === 0
                                        ? `Chapter ${si + 1} Title`
                                        : `Chapter ${si + 1} paragraph ${pi} body content for realistic document simulation.`,
                                bold: pi === 0,
                            }),
                        ],
                    });
                }),
            })),
        });
        await PackerOrig.toBuffer(doc);
    });
});
