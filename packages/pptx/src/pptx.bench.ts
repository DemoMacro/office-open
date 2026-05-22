import PptxGenJS from "pptxgenjs";
import { bench, describe } from "vite-plus/test";

import { Packer } from "./export";
import { Paragraph, Presentation, TextRun, Shape, Table } from "./index";

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

describe("PPTX: Object Creation (no pack)", () => {
  bench("ours — simple (2 shapes)", () => {
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

  bench("pptxgenjs — simple (2 shapes)", () => {
    const pptx = new PptxGenJS();
    const slide = pptx.addSlide();
    slide.addText("Hello World", { x: 1, y: 1, w: 4, h: 2 });
    slide.addText("Second shape", { x: 2, y: 3.5, w: 5, h: 1 });
  });

  bench("ours — styled shapes (20 shapes)", () => {
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
                      new TextRun({
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

  bench("pptxgenjs — styled shapes (20 shapes)", () => {
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

  bench("ours — table (10x5)", () => {
    new Presentation({
      slides: [
        {
          children: [
            new Table({
              rows: TABLE_ROWS.map((row) => ({
                cells: row.cells.map((cell) => ({
                  text: cell.text,
                  fill: cell.fill ? { type: "solid", color: cell.fill } : undefined,
                })),
              })),
            }),
          ],
        },
      ],
    });
  });

  bench("pptxgenjs — table (10x5)", () => {
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

  bench("ours — full featured", () => {
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
                    new TextRun({
                      text: "Title Slide",
                      fontSize: 28,
                      bold: true,
                    }),
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
                        new TextRun({
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
                  fill: cell.fill ? { type: "solid", color: cell.fill } : undefined,
                })),
              })),
            }),
          ],
        },
      ],
    });
  });

  bench("pptxgenjs — full featured", () => {
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

describe("PPTX: Create + toBuffer", () => {
  // NOTE: pptxgenjs write({outputType}) defaults to compression=false (JSZip STORE).
  // Our Packer.toBuffer() uses fflate level:6 (DEFLATE).
  // For fair comparison, test both pptxgenjs DEFLATE and STORE modes.

  bench("ours — simple (2 shapes) + toBuffer", async () => {
    const pres = new Presentation({
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
    await Packer.toBuffer(pres);
  });

  bench("pptxgenjs — simple (2 shapes) + write(nodebuffer, DEFLATE)", async () => {
    const pptx = new PptxGenJS();
    const slide = pptx.addSlide();
    slide.addText("Hello World", { x: 1, y: 1, w: 4, h: 2 });
    slide.addText("Second shape", { x: 2, y: 3.5, w: 5, h: 1 });
    await pptx.write({ outputType: "nodebuffer", compression: true });
  });

  bench("pptxgenjs — simple (2 shapes) + write(nodebuffer, STORE)", async () => {
    const pptx = new PptxGenJS();
    const slide = pptx.addSlide();
    slide.addText("Hello World", { x: 1, y: 1, w: 4, h: 2 });
    slide.addText("Second shape", { x: 2, y: 3.5, w: 5, h: 1 });
    await pptx.write({ outputType: "nodebuffer" });
  });

  bench("ours — styled shapes (20 shapes) + toBuffer", async () => {
    const pres = new Presentation({
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
                      new TextRun({
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
    await Packer.toBuffer(pres);
  });

  bench("pptxgenjs — styled shapes (20 shapes) + write(nodebuffer, DEFLATE)", async () => {
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
    await pptx.write({ outputType: "nodebuffer", compression: true });
  });

  bench("pptxgenjs — styled shapes (20 shapes) + write(nodebuffer, STORE)", async () => {
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
    await pptx.write({ outputType: "nodebuffer" });
  });

  bench("ours — table (10x5) + toBuffer", async () => {
    const pres = new Presentation({
      slides: [
        {
          children: [
            new Table({
              rows: TABLE_ROWS.map((row) => ({
                cells: row.cells.map((cell) => ({
                  text: cell.text,
                  fill: cell.fill ? { type: "solid", color: cell.fill } : undefined,
                })),
              })),
            }),
          ],
        },
      ],
    });
    await Packer.toBuffer(pres);
  });

  bench("pptxgenjs — table (10x5) + write(nodebuffer, DEFLATE)", async () => {
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
    await pptx.write({ outputType: "nodebuffer", compression: true });
  });

  bench("pptxgenjs — table (10x5) + write(nodebuffer, STORE)", async () => {
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
    await pptx.write({ outputType: "nodebuffer" });
  });

  bench("ours — full featured + toBuffer", async () => {
    const pres = new Presentation({
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
                    new TextRun({
                      text: "Title Slide",
                      fontSize: 28,
                      bold: true,
                    }),
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
                        new TextRun({
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
                  fill: cell.fill ? { type: "solid", color: cell.fill } : undefined,
                })),
              })),
            }),
          ],
        },
      ],
    });
    await Packer.toBuffer(pres);
  });

  bench("pptxgenjs — full featured + write(nodebuffer, DEFLATE)", async () => {
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
    await pptx.write({ outputType: "nodebuffer", compression: true });
  });

  bench("pptxgenjs — full featured + write(nodebuffer, STORE)", async () => {
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
    await pptx.write({ outputType: "nodebuffer" });
  });
});

// ── Large file benchmarks ──

const LARGE_SHAPES = Array.from({ length: 100 }, (_, i) => ({
  text: `Shape ${i + 1} text content with sample data for realistic presentation simulation.`,
  bold: i % 3 === 0,
  italic: i % 5 === 0,
  fontSize: 12 + (i % 6) * 2,
  fill: i % 4 === 0 ? "4472C4" : undefined,
}));

const LARGE_TABLE_ROWS = Array.from({ length: 50 }, (_, rowIdx) => ({
  cells: Array.from({ length: 10 }, (_, colIdx) => ({
    text: `R${rowIdx + 1}C${colIdx + 1} data`,
    fill: rowIdx === 0 ? "4472C4" : undefined,
  })),
}));

describe("PPTX: Large Files — Create + toBuffer", () => {
  bench("ours — 10 slides × 10 shapes + toBuffer", async () => {
    const pres = new Presentation({
      slides: Array.from({ length: 10 }, (_, si) => ({
        children: Array.from({ length: 10 }, (_, shi) => {
          const s = LARGE_SHAPES[si * 10 + shi];
          return new Shape({
            x: 100 + (shi % 3) * 300,
            y: 100 + Math.floor(shi / 3) * 200,
            width: 250,
            height: 80,
            fill: s.fill,
            paragraphs: [
              new Paragraph({
                properties: { bulletNone: true },
                children: [
                  new TextRun({
                    text: s.text,
                    bold: s.bold,
                    italic: s.italic,
                    fontSize: s.fontSize,
                  }),
                ],
              }),
            ],
          });
        }),
      })),
    });
    await Packer.toBuffer(pres);
  });

  bench("pptxgenjs — 10 slides × 10 shapes + write(DEFLATE)", async () => {
    const pptx = new PptxGenJS();
    for (let si = 0; si < 10; si++) {
      const slide = pptx.addSlide();
      for (let shi = 0; shi < 10; shi++) {
        const s = LARGE_SHAPES[si * 10 + shi];
        slide.addText(
          [
            {
              text: s.text,
              options: { bold: s.bold, italic: s.italic, fontSize: s.fontSize },
            },
          ],
          {
            x: 1 + (shi % 3) * 3,
            y: 1 + Math.floor(shi / 3) * 2,
            w: 2.5,
            h: 0.8,
            fill: s.fill ? { color: s.fill } : undefined,
          },
        );
      }
    }
    await pptx.write({ outputType: "nodebuffer", compression: true });
  });

  bench("ours — 50x10 table + toBuffer", async () => {
    const pres = new Presentation({
      slides: [
        {
          children: [
            new Table({
              rows: LARGE_TABLE_ROWS.map((row) => ({
                cells: row.cells.map((cell) => ({
                  text: cell.text,
                  fill: cell.fill ? { type: "solid", color: cell.fill } : undefined,
                })),
              })),
            }),
          ],
        },
      ],
    });
    await Packer.toBuffer(pres);
  });

  bench("pptxgenjs — 50x10 table + write(DEFLATE)", async () => {
    const pptx = new PptxGenJS();
    const slide = pptx.addSlide();
    slide.addTable(
      LARGE_TABLE_ROWS.map((row) =>
        row.cells.map((cell) => ({
          text: cell.text,
          options: cell.fill ? { fill: { color: cell.fill } } : undefined,
        })),
      ),
      { x: 0.5, y: 0.5, w: 9, h: 5 },
    );
    await pptx.write({ outputType: "nodebuffer", compression: true });
  });

  bench("ours — 20 slides full + toBuffer", async () => {
    const pres = new Presentation({
      slides: Array.from({ length: 20 }, (_, si) => ({
        children: [
          new Shape({
            x: 100,
            y: 50,
            width: 800,
            height: 60,
            fill: "4472C4",
            paragraphs: [
              new Paragraph({
                properties: { alignment: "CENTER", bulletNone: true },
                children: [
                  new TextRun({
                    text: `Slide ${si + 1} Title`,
                    fontSize: 28,
                    bold: true,
                  }),
                ],
              }),
            ],
          }),
          ...Array.from({ length: 5 }, (_, j) => {
            const s = LARGE_SHAPES[si * 5 + j];
            return new Shape({
              x: 100 + (j % 2) * 400,
              y: 150 + Math.floor(j / 2) * 150,
              width: 350,
              height: 100,
              fill: s.fill,
              paragraphs: [
                new Paragraph({
                  properties: { bulletNone: true },
                  children: [new TextRun({ text: s.text, bold: s.bold, fontSize: 14 })],
                }),
              ],
            });
          }),
        ],
      })),
    });
    await Packer.toBuffer(pres);
  });

  bench("pptxgenjs — 20 slides full + write(DEFLATE)", async () => {
    const pptx = new PptxGenJS();
    for (let si = 0; si < 20; si++) {
      const slide = pptx.addSlide();
      slide.addText(
        [
          {
            text: `Slide ${si + 1} Title`,
            options: { fontSize: 28, bold: true, color: "FFFFFF" },
          },
        ],
        { x: 1, y: 0.5, w: 8, h: 0.6, fill: { color: "4472C4" }, align: "center" },
      );
      for (let j = 0; j < 5; j++) {
        const s = LARGE_SHAPES[si * 5 + j];
        slide.addText([{ text: s.text, options: { bold: s.bold, fontSize: 14 } }], {
          x: 1 + (j % 2) * 4,
          y: 1.5 + Math.floor(j / 2) * 1.5,
          w: 3.5,
          h: 1,
          fill: s.fill ? { color: s.fill } : undefined,
        });
      }
    }
    await pptx.write({ outputType: "nodebuffer", compression: true });
  });
});
