import PptxGenJS from "pptxgenjs";
import { bench, describe } from "vitest";

import { Paragraph, Presentation, Run, Shape, Table } from "./index";

// ── Shared fixture data ──

const SHAPE_TEXTS = Array.from({ length: 20 }, (_, i) => ({
    text: `Shape text line ${i + 1} with some content`,
    bold: i % 3 === 0,
    italic: i % 5 === 0,
    fontSize: 14 + (i % 4) * 2,
}));

const TABLE_ROWS = Array.from({ length: 10 }, (_, rowIdx) => ({
    cells: Array.from({ length: 5 }, (_, colIdx) => ({
        text: `R${rowIdx + 1}C${colIdx + 1}`,
        fill: rowIdx === 0 ? "4472C4" : undefined,
    })),
}));

// ── Benchmarks ──

describe("Benchmark: Presentation creation — ours vs pptxgenjs", () => {
    bench("Presentation() simple (2 shapes)", () => {
        new Presentation({
            slides: [
                {
                    children: [
                        new Shape({ x: 100, y: 100, width: 400, height: 200, text: "Hello World" }),
                        new Shape({
                            x: 200,
                            y: 350,
                            width: 500,
                            height: 100,
                            text: "Second shape",
                        }),
                    ],
                },
            ],
        });
    });

    bench("PptxGenJS() simple (2 shapes)", () => {
        const pptx = new PptxGenJS();
        const slide = pptx.addSlide();
        slide.addText("Hello World", { x: 1, y: 1, w: 4, h: 2 });
        slide.addText("Second shape", { x: 2, y: 3.5, w: 5, h: 1 });
    });

    bench("Presentation() with styled shapes (20 shapes)", () => {
        new Presentation({
            slides: [
                {
                    children: SHAPE_TEXTS.map(
                        (s) =>
                            new Shape({
                                x: 100,
                                y: 100,
                                width: 400,
                                height: 200,
                                fill: s.bold ? "4472C4" : undefined,
                                paragraphs: [
                                    new Paragraph({
                                        properties: { bulletNone: true },
                                        children: [
                                            new Run({
                                                text: s.text,
                                                bold: s.bold,
                                                italic: s.italic,
                                                fontSize: s.fontSize,
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                    ),
                },
            ],
        });
    });

    bench("PptxGenJS() with styled shapes (20 shapes)", () => {
        const pptx = new PptxGenJS();
        const slide = pptx.addSlide();
        for (const s of SHAPE_TEXTS) {
            slide.addText(
                [
                    {
                        text: s.text,
                        options: { bold: s.bold, italic: s.italic, fontSize: s.fontSize },
                    },
                ],
                { x: 1, y: 1, w: 4, h: 2, fill: s.bold ? { color: "4472C4" } : undefined },
            );
        }
    });

    bench("Presentation() with table (10x5)", () => {
        new Presentation({
            slides: [
                {
                    children: [
                        new Table({
                            rows: TABLE_ROWS.map((row) => ({
                                cells: row.cells.map((cell) => ({
                                    text: cell.text,
                                    fill: cell.fill
                                        ? { type: "solid", color: cell.fill }
                                        : undefined,
                                })),
                            })),
                        }),
                    ],
                },
            ],
        });
    });

    bench("PptxGenJS() with table (10x5)", () => {
        const pptx = new PptxGenJS();
        const slide = pptx.addSlide();
        slide.addTable(
            TABLE_ROWS.map((row) =>
                row.cells.map((cell) => ({
                    text: cell.text,
                    options: cell.fill ? { fill: { color: cell.fill } } : undefined,
                })),
            ),
            { x: 1, y: 1, w: 8, h: 4 },
        );
    });

    bench("Presentation() full featured", () => {
        new Presentation({
            slides: [
                {
                    children: [
                        new Shape({
                            x: 100,
                            y: 50,
                            width: 800,
                            height: 60,
                            text: "Title Slide",
                            geometry: "rect",
                            fill: "4472C4",
                            paragraphs: [
                                new Paragraph({
                                    properties: { alignment: "CENTER", bulletNone: true },
                                    children: [
                                        new Run({ text: "Title Slide", fontSize: 28, bold: true }),
                                    ],
                                }),
                            ],
                        }),
                        ...SHAPE_TEXTS.map(
                            (s) =>
                                new Shape({
                                    x: 100,
                                    y: 100,
                                    width: 400,
                                    height: 200,
                                    fill: s.bold ? "4472C4" : undefined,
                                    paragraphs: [
                                        new Paragraph({
                                            properties: { bulletNone: true },
                                            children: [
                                                new Run({
                                                    text: s.text,
                                                    bold: s.bold,
                                                    italic: s.italic,
                                                    fontSize: s.fontSize,
                                                }),
                                            ],
                                        }),
                                    ],
                                }),
                        ),
                        new Table({
                            rows: TABLE_ROWS.map((row) => ({
                                cells: row.cells.map((cell) => ({
                                    text: cell.text,
                                    fill: cell.fill
                                        ? { type: "solid", color: cell.fill }
                                        : undefined,
                                })),
                            })),
                        }),
                    ],
                },
            ],
        });
    });

    bench("PptxGenJS() full featured", () => {
        const pptx = new PptxGenJS();
        const slide = pptx.addSlide();
        slide.addText(
            [{ text: "Title Slide", options: { fontSize: 28, bold: true, color: "FFFFFF" } }],
            { x: 1, y: 0.5, w: 8, h: 0.6, fill: { color: "4472C4" }, align: "center" },
        );
        for (const s of SHAPE_TEXTS) {
            slide.addText(
                [
                    {
                        text: s.text,
                        options: { bold: s.bold, italic: s.italic, fontSize: s.fontSize },
                    },
                ],
                { x: 1, y: 1, w: 4, h: 2, fill: s.bold ? { color: "4472C4" } : undefined },
            );
        }
        slide.addTable(
            TABLE_ROWS.map((row) =>
                row.cells.map((cell) => ({
                    text: cell.text,
                    options: cell.fill ? { fill: { color: cell.fill } } : undefined,
                })),
            ),
            { x: 1, y: 3, w: 8, h: 4 },
        );
    });
});
