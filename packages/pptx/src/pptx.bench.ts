import {
  OoxmlMimeType,
  ZIP_STORED_LEVEL,
  zipAndConvert,
  zipSyncAndConvert,
} from "@office-open/core";
import PptxGenJS from "pptxgenjs";
import { bench, describe } from "vite-plus/test";

import { Packer } from "./export";
import { Paragraph, Presentation, TextRun, Shape, Table } from "./index";

// Pack with STORE (no compression) for fair comparison with PptxGenJS.
const toBufferStore = (pres: Presentation) => {
  const files = Packer.compile(pres);
  return zipSyncAndConvert(files, "nodebuffer", OoxmlMimeType.PPTX, ZIP_STORED_LEVEL);
};

// Pack with STORE async (no compression, non-blocking).
const toBufferStoreAsync = async (pres: Presentation) => {
  const files = Packer.compile(pres);
  return await zipAndConvert(files, "nodebuffer", OoxmlMimeType.PPTX, ZIP_STORED_LEVEL);
};

// Helper to pack with DEFLATE or STORE sync, avoiding presentation construction duplication.
const packOurs = (pres: Presentation, store: boolean) =>
  store ? toBufferStore(pres) : Packer.toBufferSync(pres);

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

// ── Fixture helpers (ours) ──

const buildSimplePres = () =>
  new Presentation({
    slides: [
      {
        children: [
          new Shape({
            x: 100,
            y: 100,
            width: 400,
            height: 200,
            textBody: { text: "Hello World" },
          }),
          new Shape({
            x: 200,
            y: 350,
            width: 500,
            height: 100,
            textBody: { text: "Second shape" },
          }),
        ],
      },
    ],
  });

const buildStyledPres = () =>
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
              textBody: {
                children: [
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
              },
            }),
        ),
      },
    ],
  });

const buildTablePres = () =>
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

const buildFullFeaturedPres = () =>
  new Presentation({
    slides: [
      {
        children: [
          new Shape({
            x: 100,
            y: 50,
            width: 800,
            height: 60,
            textBody: {
              text: "Title Slide",
              children: [
                new Paragraph({
                  properties: { alignment: "center", bullet: { type: "none" } },
                  children: [
                    new TextRun({
                      text: "Title Slide",
                      fontSize: 28,
                      bold: true,
                    }),
                  ],
                }),
              ],
            },
            geometry: "rect",
            fill: "4472C4",
          }),
          ...SHAPE_TEXTS.map(
            (s) =>
              new Shape({
                x: 100,
                y: 100,
                width: 400,
                height: 200,
                fill: s.bold ? "4472C4" : undefined,
                textBody: {
                  children: [
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
                },
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

// ── Benchmarks ──

// Our Packer uses fflate async zip() (Web Workers) with DEFLATE by default.
// PptxGenJS always uses JSZip STORE (no compression) regardless of its compression flag.
// STORE benchmarks provide a fair apples-to-apples ZIP engine comparison.
describe("PPTX: Create + toBuffer", () => {
  bench(
    "ours DEFLATE sync — simple (2 shapes) + toBufferSync",
    () => {
      Packer.toBufferSync(buildSimplePres());
    },
    { iterations: 50 },
  );

  bench(
    "ours STORE sync — simple (2 shapes) + toBuffer",
    () => {
      toBufferStore(buildSimplePres());
    },
    { iterations: 50 },
  );

  bench(
    "ours DEFLATE async — simple (2 shapes) + toBuffer",
    async () => {
      await Packer.toBuffer(buildSimplePres());
    },
    { iterations: 50 },
  );

  bench(
    "ours STORE async — simple (2 shapes) + toBuffer",
    async () => {
      await toBufferStoreAsync(buildSimplePres());
    },
    { iterations: 50 },
  );

  bench(
    "pptxgenjs — simple (2 shapes) + toBuffer",
    async () => {
      const pptx = new PptxGenJS();
      const slide = pptx.addSlide();
      slide.addText("Hello World", { x: 1, y: 1, w: 4, h: 2 });
      slide.addText("Second shape", { x: 2, y: 3.5, w: 5, h: 1 });
      await pptx.write({ outputType: "nodebuffer", compression: true });
    },
    { iterations: 50 },
  );

  bench(
    "ours DEFLATE sync — styled shapes (20) + toBufferSync",
    () => {
      Packer.toBufferSync(buildStyledPres());
    },
    { iterations: 50 },
  );

  bench(
    "ours STORE sync — styled shapes (20) + toBuffer",
    () => {
      toBufferStore(buildStyledPres());
    },
    { iterations: 50 },
  );

  bench(
    "ours DEFLATE async — styled shapes (20) + toBuffer",
    async () => {
      await Packer.toBuffer(buildStyledPres());
    },
    { iterations: 50 },
  );

  bench(
    "ours STORE async — styled shapes (20) + toBuffer",
    async () => {
      await toBufferStoreAsync(buildStyledPres());
    },
    { iterations: 50 },
  );

  bench(
    "pptxgenjs — styled shapes (20) + toBuffer",
    async () => {
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
    },
    { iterations: 50 },
  );

  bench(
    "ours DEFLATE sync — table (10x5) + toBufferSync",
    () => {
      Packer.toBufferSync(buildTablePres());
    },
    { iterations: 50 },
  );

  bench(
    "ours STORE sync — table (10x5) + toBuffer",
    () => {
      toBufferStore(buildTablePres());
    },
    { iterations: 50 },
  );

  bench(
    "ours DEFLATE async — table (10x5) + toBuffer",
    async () => {
      await Packer.toBuffer(buildTablePres());
    },
    { iterations: 50 },
  );

  bench(
    "ours STORE async — table (10x5) + toBuffer",
    async () => {
      await toBufferStoreAsync(buildTablePres());
    },
    { iterations: 50 },
  );

  bench(
    "pptxgenjs — table (10x5) + toBuffer",
    async () => {
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
    },
    { iterations: 50 },
  );

  bench(
    "ours DEFLATE sync — full featured + toBufferSync",
    () => {
      Packer.toBufferSync(buildFullFeaturedPres());
    },
    { iterations: 50 },
  );

  bench(
    "ours STORE sync — full featured + toBuffer",
    () => {
      toBufferStore(buildFullFeaturedPres());
    },
    { iterations: 50 },
  );

  bench(
    "ours DEFLATE async — full featured + toBuffer",
    async () => {
      await Packer.toBuffer(buildFullFeaturedPres());
    },
    { iterations: 50 },
  );

  bench(
    "ours STORE async — full featured + toBuffer",
    async () => {
      await toBufferStoreAsync(buildFullFeaturedPres());
    },
    { iterations: 50 },
  );

  bench(
    "pptxgenjs — full featured + toBuffer",
    async () => {
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
    },
    { iterations: 50 },
  );
});

// ── Large file benchmarks ──

const LARGE_SHAPES = Array.from({ length: 600 }, (_, i) => ({
  text: `Shape ${i + 1} text content with sample data for realistic presentation simulation and stress testing.`,
  bold: i % 3 === 0,
  italic: i % 5 === 0,
  fontSize: 12 + (i % 6) * 2,
  fill: i % 4 === 0 ? "4472C4" : undefined,
}));

const LARGE_TABLE_ROWS = Array.from({ length: 100 }, (_, rowIdx) => ({
  cells: Array.from({ length: 10 }, (_, colIdx) => ({
    text: `R${rowIdx + 1}C${colIdx + 1} data content for stress test`,
    fill: rowIdx === 0 ? "4472C4" : undefined,
  })),
}));

const build30Slides20Shapes = () =>
  new Presentation({
    slides: Array.from({ length: 30 }, (_, si) => ({
      children: Array.from({ length: 20 }, (_, shi) => {
        const s = LARGE_SHAPES[(si * 20 + shi) % LARGE_SHAPES.length];
        return new Shape({
          x: 50 + (shi % 4) * 230,
          y: 50 + Math.floor(shi / 4) * 110,
          width: 200,
          height: 80,
          fill: s.fill,
          textBody: {
            children: [
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
          },
        });
      }),
    })),
  });

const build100x10Table = () =>
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

const build50SlidesFull = () =>
  new Presentation({
    slides: Array.from({ length: 50 }, (_, si) => ({
      children: [
        new Shape({
          x: 100,
          y: 50,
          width: 800,
          height: 60,
          fill: "4472C4",
          textBody: {
            children: [
              new Paragraph({
                properties: { alignment: "center", bullet: { type: "none" } },
                children: [
                  new TextRun({ text: `Slide ${si + 1} Title`, fontSize: 28, bold: true }),
                ],
              }),
            ],
          },
        }),
        ...Array.from({ length: 10 }, (_, j) => {
          const s = LARGE_SHAPES[(si * 10 + j) % LARGE_SHAPES.length];
          return new Shape({
            x: 50 + (j % 3) * 300,
            y: 150 + Math.floor(j / 3) * 120,
            width: 270,
            height: 90,
            fill: s.fill,
            textBody: {
              children: [
                new Paragraph({
                  properties: { bullet: { type: "none" } },
                  children: [
                    new TextRun({ text: s.text, bold: s.bold, italic: s.italic, fontSize: 14 }),
                  ],
                }),
              ],
            },
          });
        }),
      ],
    })),
  });

describe("PPTX: Large Files — Create + toBuffer", () => {
  bench(
    "ours DEFLATE sync — 30 slides × 20 shapes + toBufferSync",
    () => {
      packOurs(build30Slides20Shapes(), false);
    },
    { iterations: 10 },
  );

  bench(
    "ours STORE sync — 30 slides × 20 shapes + toBuffer",
    () => {
      packOurs(build30Slides20Shapes(), true);
    },
    { iterations: 10 },
  );

  bench(
    "ours DEFLATE async — 30 slides × 20 shapes + toBuffer",
    async () => {
      await Packer.toBuffer(build30Slides20Shapes());
    },
    { iterations: 10 },
  );

  bench(
    "ours STORE async — 30 slides × 20 shapes + toBuffer",
    async () => {
      await toBufferStoreAsync(build30Slides20Shapes());
    },
    { iterations: 10 },
  );

  bench(
    "pptxgenjs — 30 slides × 20 shapes + toBuffer",
    async () => {
      const pptx = new PptxGenJS();
      for (let si = 0; si < 30; si++) {
        const slide = pptx.addSlide();
        for (let shi = 0; shi < 20; shi++) {
          const s = LARGE_SHAPES[(si * 20 + shi) % LARGE_SHAPES.length];
          slide.addText(
            [{ text: s.text, options: { bold: s.bold, italic: s.italic, fontSize: s.fontSize } }],
            {
              x: 0.5 + (shi % 4) * 2.3,
              y: 0.5 + Math.floor(shi / 4) * 1.1,
              w: 2,
              h: 0.8,
              fill: s.fill ? { color: s.fill } : undefined,
            },
          );
        }
      }
      await pptx.write({ outputType: "nodebuffer", compression: true });
    },
    { iterations: 10 },
  );

  bench(
    "ours DEFLATE sync — 100x10 table + toBufferSync",
    () => {
      packOurs(build100x10Table(), false);
    },
    { iterations: 10 },
  );

  bench(
    "ours STORE sync — 100x10 table + toBuffer",
    () => {
      packOurs(build100x10Table(), true);
    },
    { iterations: 10 },
  );

  bench(
    "ours DEFLATE async — 100x10 table + toBuffer",
    async () => {
      await Packer.toBuffer(build100x10Table());
    },
    { iterations: 10 },
  );

  bench(
    "ours STORE async — 100x10 table + toBuffer",
    async () => {
      await toBufferStoreAsync(build100x10Table());
    },
    { iterations: 10 },
  );

  bench(
    "pptxgenjs — 100x10 table + toBuffer",
    async () => {
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
    },
    { iterations: 10 },
  );

  bench(
    "ours DEFLATE sync — 50 slides full + toBufferSync",
    () => {
      packOurs(build50SlidesFull(), false);
    },
    { iterations: 10 },
  );

  bench(
    "ours STORE sync — 50 slides full + toBuffer",
    () => {
      packOurs(build50SlidesFull(), true);
    },
    { iterations: 10 },
  );

  bench(
    "ours DEFLATE async — 50 slides full + toBuffer",
    async () => {
      await Packer.toBuffer(build50SlidesFull());
    },
    { iterations: 10 },
  );

  bench(
    "ours STORE async — 50 slides full + toBuffer",
    async () => {
      await toBufferStoreAsync(build50SlidesFull());
    },
    { iterations: 10 },
  );

  bench(
    "pptxgenjs — 50 slides full + toBuffer",
    async () => {
      const pptx = new PptxGenJS();
      for (let si = 0; si < 50; si++) {
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
        for (let j = 0; j < 10; j++) {
          const s = LARGE_SHAPES[(si * 10 + j) % LARGE_SHAPES.length];
          slide.addText(
            [{ text: s.text, options: { bold: s.bold, italic: s.italic, fontSize: 14 } }],
            {
              x: 0.5 + (j % 3) * 3,
              y: 1.5 + Math.floor(j / 3) * 1.2,
              w: 2.7,
              h: 0.9,
              fill: s.fill ? { color: s.fill } : undefined,
            },
          );
        }
      }
      await pptx.write({ outputType: "nodebuffer", compression: true });
    },
    { iterations: 10 },
  );
});

// ── Large file ~100MB mixed benchmarks ──

// Generate a unique 500KB fake image per seed — different seeds produce different
// byte patterns, so images have unique hashes and won't be deduplicated by the packer.
const makeImage = (seed: number): Uint8Array => {
  const size = 500 * 1024;
  const buf = new Uint8Array(size);
  for (let i = 0; i < size; i++) buf[i] = (i * 7 + seed * 13 + 37) & 0xff;
  return buf;
};

const MIXED_IMAGES = Array.from({ length: 200 }, (_, i) => makeImage(i));

const MIXED_IMAGES_BASE64 = MIXED_IMAGES.map(
  (img) => `data:image/jpeg;base64,${Buffer.from(img).toString("base64")}`,
);

const buildMixed100MbPres = () =>
  new Presentation({
    slides: Array.from({ length: 100 }, (_, si) => ({
      children: [
        // 5 shapes
        ...Array.from({ length: 5 }, (_, shi) => {
          const s = LARGE_SHAPES[(si * 5 + shi) % LARGE_SHAPES.length];
          return new Shape({
            x: 50 + shi * 180,
            y: 50,
            width: 160,
            height: 80,
            fill: s.fill,
            textBody: {
              children: [
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
            },
          });
        }),
        // 2 unique images (500KB each)
        ...Array.from({ length: 2 }, (_, pi) => ({
          picture: {
            x: 50 + pi * 350,
            y: 200,
            width: 300,
            height: 200,
            data: MIXED_IMAGES[(si * 2 + pi) % MIXED_IMAGES.length],
            type: "jpg" as const,
          },
        })),
        // 5×5 table
        new Table({
          rows: Array.from({ length: 5 }, (_, rowIdx) => ({
            cells: Array.from({ length: 5 }, (_, colIdx) => ({
              text: `S${si + 1}R${rowIdx + 1}C${colIdx + 1}`,
              fill: rowIdx === 0 ? "4472C4" : undefined,
            })),
          })),
        }),
      ],
    })),
  });

describe("PPTX: Large File (~100MB) — Mixed + async vs sync", () => {
  bench(
    "ours DEFLATE sync — mixed (100sl) + toBufferSync",
    () => {
      Packer.toBufferSync(buildMixed100MbPres());
    },
    { iterations: 3 },
  );

  bench(
    "ours STORE sync — mixed (100sl) + toBufferStore",
    () => {
      toBufferStore(buildMixed100MbPres());
    },
    { iterations: 3 },
  );

  bench(
    "ours DEFLATE async — mixed (100sl) + toBuffer",
    async () => {
      await Packer.toBuffer(buildMixed100MbPres());
    },
    { iterations: 3 },
  );

  bench(
    "ours STORE async — mixed (100sl) + toBuffer",
    async () => {
      await toBufferStoreAsync(buildMixed100MbPres());
    },
    { iterations: 3 },
  );

  bench(
    "pptxgenjs — mixed (100sl) + toBuffer",
    async () => {
      const pptx = new PptxGenJS();
      for (let si = 0; si < 100; si++) {
        const slide = pptx.addSlide();
        // 5 shapes
        for (let shi = 0; shi < 5; shi++) {
          const s = LARGE_SHAPES[(si * 5 + shi) % LARGE_SHAPES.length];
          slide.addText(
            [{ text: s.text, options: { bold: s.bold, italic: s.italic, fontSize: s.fontSize } }],
            {
              x: 0.5 + shi * 1.8,
              y: 0.5,
              w: 1.6,
              h: 0.8,
              fill: s.fill ? { color: s.fill } : undefined,
            },
          );
        }
        // 2 images
        for (let pi = 0; pi < 2; pi++) {
          slide.addImage({
            data: MIXED_IMAGES_BASE64[(si * 2 + pi) % MIXED_IMAGES_BASE64.length],
            x: 0.5 + pi * 3.5,
            y: 2,
            w: 3,
            h: 2,
          });
        }
        // 5×5 table
        slide.addTable(
          Array.from({ length: 5 }, (_, rowIdx) =>
            Array.from({ length: 5 }, (_, colIdx) => ({
              text: `S${si + 1}R${rowIdx + 1}C${colIdx + 1}`,
              options: rowIdx === 0 ? { fill: { color: "4472C4" } } : undefined,
            })),
          ),
          { x: 0.5, y: 4.5, w: 9, h: 2 },
        );
      }
      await pptx.write({ outputType: "nodebuffer", compression: true });
    },
    { iterations: 3 },
  );
});
