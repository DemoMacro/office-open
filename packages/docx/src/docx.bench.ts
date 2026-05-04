import {
    AlignmentType as AlignmentTypeOrig,
    Document as DocumentOrig,
    Footer as FooterOrig,
    Header as HeaderOrig,
    HeadingLevel as HeadingLevelOrig,
    PageNumber as PageNumberOrig,
    Paragraph as ParagraphOrig,
    Table as TableOrig,
    TableCell as TableCellOrig,
    TableRow as TableRowOrig,
    TextRun as TextRunOrig,
} from "docx";
import { bench, describe } from "vite-plus/test";

import {
    AlignmentType,
    Document,
    Footer,
    Header,
    HeadingLevel,
    PageNumber,
    Paragraph,
    Table,
    TableCell,
    TableRow,
    TextRun,
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

describe("Benchmark: Document creation — ours vs original docx", () => {
    bench("Document() simple", () => {
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

    bench("Document (original) simple", () => {
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

    bench("Document() with styled paragraphs", () => {
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

    bench("Document (original) with styled paragraphs", () => {
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

    bench("Document() with table", () => {
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

    bench("Document (original) with table", () => {
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

    bench("Document() full featured", () => {
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

    bench("Document (original) full featured", () => {
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
