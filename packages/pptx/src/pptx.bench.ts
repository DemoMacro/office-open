import { OoxmlMimeType, ZIP_STORED_LEVEL, zipAndConvert } from "@office-open/core";
import PptxGenJS from "pptxgenjs";
import { bench, describe } from "vite-plus/test";

import { Packer } from "./export";
import { Paragraph, Presentation, TextRun, Shape, Table } from "./index";

// Pack with STORE (no compression) for fair comparison with PptxGenJS.
const toBufferStore = (pres: Presentation) => {
  const files = Packer.compile(pres);
  return zipAndConvert(files, "nodebuffer", OoxmlMimeType.PPTX, ZIP_STORED_LEVEL);
};

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
                    properties: { bullet: { type: "none" } },
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
                  properties: { alignment: "CENTER", bullet: { type: "none" } },
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
                      properties: { bullet: { type: "none" } },
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

// Our Packer uses fflate async zip() (Web Workers) with DEFLATE by default.
// PptxGenJS always uses JSZip STORE (no compression) regardless of its compression flag.
// STORE benchmarks provide a fair apples-to-apples ZIP engine comparison.
describe("PPTX: Create + toBuffer", () => {
  bench("ours DEFLATE — simple (2 shapes) + toBuffer", async () => {
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

  bench("ours STORE — simple (2 shapes) + toBuffer", async () => {
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
    await toBufferStore(pres);
  });

  bench("pptxgenjs — simple (2 shapes) + toBuffer", async () => {
    const pptx = new PptxGenJS();
    const slide = pptx.addSlide();
    slide.addText("Hello World", { x: 1, y: 1, w: 4, h: 2 });
    slide.addText("Second shape", { x: 2, y: 3.5, w: 5, h: 1 });
    await pptx.write({ outputType: "nodebuffer", compression: true });
  });

  bench("ours DEFLATE — styled shapes (20 shapes) + toBuffer", async () => {
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
                    properties: { bullet: { type: "none" } },
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

  bench("ours STORE — styled shapes (20 shapes) + toBuffer", async () => {
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
                    properties: { bullet: { type: "none" } },
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
    await toBufferStore(pres);
  });

  bench("pptxgenjs — styled shapes (20 shapes) + toBuffer", async () => {
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

  bench("ours DEFLATE — table (10x5) + toBuffer", async () => {
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

  bench("ours STORE — table (10x5) + toBuffer", async () => {
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
    await toBufferStore(pres);
  });

  bench("pptxgenjs — table (10x5) + toBuffer", async () => {
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

  bench("ours DEFLATE — full featured + toBuffer", async () => {
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
                  properties: { alignment: "CENTER", bullet: { type: "none" } },
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
                      properties: { bullet: { type: "none" } },
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

  bench("ours STORE — full featured + toBuffer", async () => {
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
                  properties: { alignment: "CENTER", bullet: { type: "none" } },
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
                      properties: { bullet: { type: "none" } },
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
    await toBufferStore(pres);
  });

  bench("pptxgenjs — full featured + toBuffer", async () => {
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

// Helper to pack with DEFLATE or STORE, avoiding presentation construction duplication.
const packOurs = (pres: Presentation, store: boolean) =>
  store ? toBufferStore(pres) : Packer.toBuffer(pres);

const build10Slides10Shapes = () =>
  new Presentation({
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
              properties: { bullet: { type: "none" } },
              children: [
                new TextRun({ text: s.text, bold: s.bold, italic: s.italic, fontSize: s.fontSize }),
              ],
            }),
          ],
        });
      }),
    })),
  });

const build50x10Table = () =>
  new Presentation({
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

const build20SlidesFull = () =>
  new Presentation({
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
              properties: { alignment: "CENTER", bullet: { type: "none" } },
              children: [new TextRun({ text: `Slide ${si + 1} Title`, fontSize: 28, bold: true })],
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
                properties: { bullet: { type: "none" } },
                children: [new TextRun({ text: s.text, bold: s.bold, fontSize: 14 })],
              }),
            ],
          });
        }),
      ],
    })),
  });

describe("PPTX: Large Files — Create + toBuffer", () => {
  bench("ours DEFLATE — 10 slides × 10 shapes + toBuffer", async () => {
    await packOurs(build10Slides10Shapes(), false);
  });

  bench("ours STORE — 10 slides × 10 shapes + toBuffer", async () => {
    await packOurs(build10Slides10Shapes(), true);
  });

  bench("pptxgenjs — 10 slides × 10 shapes + toBuffer", async () => {
    const pptx = new PptxGenJS();
    for (let si = 0; si < 10; si++) {
      const slide = pptx.addSlide();
      for (let shi = 0; shi < 10; shi++) {
        const s = LARGE_SHAPES[si * 10 + shi];
        slide.addText(
          [{ text: s.text, options: { bold: s.bold, italic: s.italic, fontSize: s.fontSize } }],
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

  bench("ours DEFLATE — 50x10 table + toBuffer", async () => {
    await packOurs(build50x10Table(), false);
  });

  bench("ours STORE — 50x10 table + toBuffer", async () => {
    await packOurs(build50x10Table(), true);
  });

  bench("pptxgenjs — 50x10 table + toBuffer", async () => {
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

  bench("ours DEFLATE — 20 slides full + toBuffer", async () => {
    await packOurs(build20SlidesFull(), false);
  });

  bench("ours STORE — 20 slides full + toBuffer", async () => {
    await packOurs(build20SlidesFull(), true);
  });

  bench("pptxgenjs — 20 slides full + toBuffer", async () => {
    const pptx = new PptxGenJS();
    for (let si = 0; si < 20; si++) {
      const slide = pptx.addSlide();
      slide.addText(
        [{ text: `Slide ${si + 1} Title`, options: { fontSize: 28, bold: true, color: "FFFFFF" } }],
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
