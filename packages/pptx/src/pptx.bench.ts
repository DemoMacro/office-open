import PptxGenJS from "pptxgenjs";
import { bench, describe } from "vite-plus/test";

import { generate, generateSync } from "./export";
import type { PresentationOptions, SlideChild } from "./file";

// Bench modes:
//   "ours default"  = XML DEFLATE level 1 (SuperFast, MS Office), media STORE — no options passed.
//   "ours all-store" = all entries STORE — { compression: { xml: 0 } }.
//
// PptxGenJS (JSZip): async only. Default STORE; DEFLATE via compression: true.
// Global compression — ALL entries including images get same treatment.

// ── Image generation ──

const makeImage = (seed: number, sizeKB: number): Uint8Array => {
  const size = sizeKB * 1024;
  const buf = new Uint8Array(size);
  for (let i = 0; i < size; i++) buf[i] = (i * 7 + seed * 13 + 37) & 0xff;
  return buf;
};

const SMALL_IMAGES = Array.from({ length: 3 }, (_, i) => makeImage(i, 200));
const SMALL_IMAGES_BASE64 = SMALL_IMAGES.map(
  (img) => `data:image/jpeg;base64,${Buffer.from(img).toString("base64")}`,
);
const LARGE_IMAGES = Array.from({ length: 20 }, (_, i) => makeImage(i, 500));
const LARGE_IMAGES_BASE64 = LARGE_IMAGES.map(
  (img) => `data:image/jpeg;base64,${Buffer.from(img).toString("base64")}`,
);

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

const buildSimplePres = (): PresentationOptions => ({
  slides: [
    {
      children: [
        {
          shape: {
            x: 100,
            y: 100,
            width: 400,
            height: 200,
            textBody: { text: "Hello World" },
          },
        },
        {
          shape: {
            x: 200,
            y: 350,
            width: 500,
            height: 100,
            textBody: { text: "Second shape" },
          },
        },
        {
          picture: {
            x: 600,
            y: 100,
            width: 300,
            height: 200,
            data: SMALL_IMAGES[0],
            type: "jpg" as const,
          },
        },
      ],
    },
  ],
});

const buildStyledPres = (): PresentationOptions => ({
  slides: [
    {
      children: [
        ...SHAPE_TEXTS.map(
          (s) =>
            ({
              shape: {
                x: 100,
                y: 100,
                width: 400,
                height: 200,
                fill: s.bold ? "4472C4" : undefined,
                textBody: {
                  children: [
                    {
                      properties: { bullet: { type: "none" } },
                      children: [
                        {
                          text: s.text,
                          bold: s.bold,
                          italic: s.italic,
                          fontSize: s.fontSize,
                        },
                      ],
                    },
                  ],
                },
              },
            }) as const,
        ),
        {
          picture: {
            x: 600,
            y: 100,
            width: 300,
            height: 200,
            data: SMALL_IMAGES[1],
            type: "jpg" as const,
          },
        },
      ],
    },
  ],
});

const buildTablePres = (): PresentationOptions => ({
  slides: [
    {
      children: [
        {
          table: {
            rows: TABLE_ROWS.map((row) => ({
              cells: row.cells.map((cell) => ({
                text: cell.text,
                fill: cell.fill ? { type: "solid", color: cell.fill } : undefined,
              })),
            })),
          },
        },
      ],
    },
  ],
});

const buildFullFeaturedPres = (): PresentationOptions => ({
  slides: [
    {
      children: [
        {
          shape: {
            x: 100,
            y: 50,
            width: 800,
            height: 60,
            textBody: {
              text: "Title Slide",
              children: [
                {
                  properties: { alignment: "center", bullet: { type: "none" } },
                  children: [
                    {
                      text: "Title Slide",
                      fontSize: 28,
                      bold: true,
                    },
                  ],
                },
              ],
            },
            geometry: "rect",
            fill: "4472C4",
          },
        },
        ...SHAPE_TEXTS.map(
          (s) =>
            ({
              shape: {
                x: 100,
                y: 100,
                width: 400,
                height: 200,
                fill: s.bold ? "4472C4" : undefined,
                textBody: {
                  children: [
                    {
                      properties: { bullet: { type: "none" } },
                      children: [
                        {
                          text: s.text,
                          bold: s.bold,
                          italic: s.italic,
                          fontSize: s.fontSize,
                        },
                      ],
                    },
                  ],
                },
              },
            }) as const,
        ),
        {
          picture: {
            x: 600,
            y: 100,
            width: 300,
            height: 200,
            data: SMALL_IMAGES[0],
            type: "jpg" as const,
          },
        },
        {
          picture: {
            x: 600,
            y: 350,
            width: 300,
            height: 200,
            data: SMALL_IMAGES[2],
            type: "jpg" as const,
          },
        },
        {
          table: {
            rows: TABLE_ROWS.map((row) => ({
              cells: row.cells.map((cell) => ({
                text: cell.text,
                fill: cell.fill ? { type: "solid", color: cell.fill } : undefined,
              })),
            })),
          },
        },
      ],
    },
  ],
});

// ── Benchmarks ──

describe("PPTX: Create + toBuffer", () => {
  bench(
    "ours default sync — simple (2 shapes + 1 img) + toBufferSync",
    () => {
      generateSync(buildSimplePres());
    },
    { iterations: 50 },
  );

  bench(
    "ours all-store sync — simple (2 shapes + 1 img) + toBuffer",
    () => {
      generateSync(buildSimplePres(), "nodebuffer", { compression: { xml: 0 } });
    },
    { iterations: 50 },
  );

  bench(
    "ours default async — simple (2 shapes + 1 img) + toBuffer",
    async () => {
      await generate(buildSimplePres());
    },
    { iterations: 50 },
  );

  bench(
    "ours all-store async — simple (2 shapes + 1 img) + toBuffer",
    async () => {
      await generate(buildSimplePres(), "nodebuffer", { compression: { xml: 0 } });
    },
    { iterations: 50 },
  );

  bench(
    "pptxgenjs DEFLATE — simple (2 shapes + 1 img) + toBuffer (async)",
    async () => {
      const pptx = new PptxGenJS();
      const slide = pptx.addSlide();
      slide.addText("Hello World", { x: 1, y: 1, w: 4, h: 2 });
      slide.addText("Second shape", { x: 2, y: 3.5, w: 5, h: 1 });
      slide.addImage({ data: SMALL_IMAGES_BASE64[0], x: 6, y: 1, w: 3, h: 2 });
      await pptx.write({ outputType: "nodebuffer", compression: true });
    },
    { iterations: 50 },
  );

  bench(
    "pptxgenjs STORE — simple (2 shapes + 1 img) + toBuffer (async)",
    async () => {
      const pptx = new PptxGenJS();
      const slide = pptx.addSlide();
      slide.addText("Hello World", { x: 1, y: 1, w: 4, h: 2 });
      slide.addText("Second shape", { x: 2, y: 3.5, w: 5, h: 1 });
      slide.addImage({ data: SMALL_IMAGES_BASE64[0], x: 6, y: 1, w: 3, h: 2 });
      await pptx.write({ outputType: "nodebuffer" });
    },
    { iterations: 50 },
  );

  bench(
    "ours default sync — styled shapes (20) + 1 img + toBufferSync",
    () => {
      generateSync(buildStyledPres());
    },
    { iterations: 50 },
  );

  bench(
    "ours all-store sync — styled shapes (20) + 1 img + toBuffer",
    () => {
      generateSync(buildStyledPres(), "nodebuffer", { compression: { xml: 0 } });
    },
    { iterations: 50 },
  );

  bench(
    "ours default async — styled shapes (20) + 1 img + toBuffer",
    async () => {
      await generate(buildStyledPres());
    },
    { iterations: 50 },
  );

  bench(
    "ours all-store async — styled shapes (20) + 1 img + toBuffer",
    async () => {
      await generate(buildStyledPres(), "nodebuffer", { compression: { xml: 0 } });
    },
    { iterations: 50 },
  );

  bench(
    "pptxgenjs DEFLATE — styled shapes (20) + 1 img + toBuffer (async)",
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
      slide.addImage({ data: SMALL_IMAGES_BASE64[1], x: 6, y: 1, w: 3, h: 2 });
      await pptx.write({ outputType: "nodebuffer", compression: true });
    },
    { iterations: 50 },
  );

  bench(
    "pptxgenjs STORE — styled shapes (20) + 1 img + toBuffer (async)",
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
      slide.addImage({ data: SMALL_IMAGES_BASE64[1], x: 6, y: 1, w: 3, h: 2 });
      await pptx.write({ outputType: "nodebuffer" });
    },
    { iterations: 50 },
  );

  bench(
    "ours default sync — table (10x5) + toBufferSync",
    () => {
      generateSync(buildTablePres());
    },
    { iterations: 50 },
  );

  bench(
    "ours all-store sync — table (10x5) + toBuffer",
    () => {
      generateSync(buildTablePres(), "nodebuffer", { compression: { xml: 0 } });
    },
    { iterations: 50 },
  );

  bench(
    "ours default async — table (10x5) + toBuffer",
    async () => {
      await generate(buildTablePres());
    },
    { iterations: 50 },
  );

  bench(
    "ours all-store async — table (10x5) + toBuffer",
    async () => {
      await generate(buildTablePres(), "nodebuffer", { compression: { xml: 0 } });
    },
    { iterations: 50 },
  );

  bench(
    "pptxgenjs DEFLATE — table (10x5) + toBuffer",
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
    "pptxgenjs STORE — table (10x5) + toBuffer",
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
      await pptx.write({ outputType: "nodebuffer" });
    },
    { iterations: 50 },
  );

  bench(
    "ours default sync — full featured + 2 imgs + toBufferSync",
    () => {
      generateSync(buildFullFeaturedPres());
    },
    { iterations: 50 },
  );

  bench(
    "ours all-store sync — full featured + 2 imgs + toBuffer",
    () => {
      generateSync(buildFullFeaturedPres(), "nodebuffer", { compression: { xml: 0 } });
    },
    { iterations: 50 },
  );

  bench(
    "ours default async — full featured + 2 imgs + toBuffer",
    async () => {
      await generate(buildFullFeaturedPres());
    },
    { iterations: 50 },
  );

  bench(
    "ours all-store async — full featured + 2 imgs + toBuffer",
    async () => {
      await generate(buildFullFeaturedPres(), "nodebuffer", { compression: { xml: 0 } });
    },
    { iterations: 50 },
  );

  bench(
    "pptxgenjs DEFLATE — full featured + 2 imgs + toBuffer (async)",
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
      slide.addImage({ data: SMALL_IMAGES_BASE64[0], x: 6, y: 1, w: 3, h: 2 });
      slide.addImage({ data: SMALL_IMAGES_BASE64[2], x: 6, y: 3.5, w: 3, h: 2 });
      slide.addTable(
        TABLE_ROWS.map((row) =>
          row.cells.map((cell) => ({
            text: cell.text,
            options: cell.fill ? { fill: { color: cell.fill } } : undefined,
          })),
        ),
        { x: 1, y: 6, w: 8, h: 4 },
      );
      await pptx.write({ outputType: "nodebuffer", compression: true });
    },
    { iterations: 50 },
  );

  bench(
    "pptxgenjs STORE — full featured + 2 imgs + toBuffer (async)",
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
      slide.addImage({ data: SMALL_IMAGES_BASE64[0], x: 6, y: 1, w: 3, h: 2 });
      slide.addImage({ data: SMALL_IMAGES_BASE64[2], x: 6, y: 3.5, w: 3, h: 2 });
      slide.addTable(
        TABLE_ROWS.map((row) =>
          row.cells.map((cell) => ({
            text: cell.text,
            options: cell.fill ? { fill: { color: cell.fill } } : undefined,
          })),
        ),
        { x: 1, y: 6, w: 8, h: 4 },
      );
      await pptx.write({ outputType: "nodebuffer" });
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

const build30Slides20Shapes = (): PresentationOptions => ({
  slides: Array.from({ length: 30 }, (_, si) => ({
    children: Array.from({ length: 20 }, (_, shi) => {
      const s = LARGE_SHAPES[(si * 20 + shi) % LARGE_SHAPES.length];
      return {
        shape: {
          x: 50 + (shi % 4) * 230,
          y: 50 + Math.floor(shi / 4) * 110,
          width: 200,
          height: 80,
          fill: s.fill,
          textBody: {
            children: [
              {
                properties: { bullet: { type: "none" } },
                children: [
                  {
                    text: s.text,
                    bold: s.bold,
                    italic: s.italic,
                    fontSize: s.fontSize,
                  },
                ],
              },
            ],
          },
        },
      };
    }),
  })),
});

const build100x10Table = (): PresentationOptions => ({
  slides: [
    {
      children: [
        {
          table: {
            rows: LARGE_TABLE_ROWS.map((row) => ({
              cells: row.cells.map((cell) => ({
                text: cell.text,
                fill: cell.fill ? { type: "solid", color: cell.fill } : undefined,
              })),
            })),
          },
        },
      ],
    },
  ],
});

const build50SlidesFull = (): PresentationOptions => ({
  slides: Array.from({ length: 50 }, (_, si) => ({
    children: [
      {
        shape: {
          x: 100,
          y: 50,
          width: 800,
          height: 60,
          fill: "4472C4",
          textBody: {
            children: [
              {
                properties: { alignment: "center", bullet: { type: "none" } },
                children: [{ text: `Slide ${si + 1} Title`, fontSize: 28, bold: true }],
              },
            ],
          },
        },
      },
      ...Array.from({ length: 10 }, (_, j) => {
        const s = LARGE_SHAPES[(si * 10 + j) % LARGE_SHAPES.length];
        return {
          shape: {
            x: 50 + (j % 3) * 300,
            y: 150 + Math.floor(j / 3) * 120,
            width: 270,
            height: 90,
            fill: s.fill,
            textBody: {
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [{ text: s.text, bold: s.bold, italic: s.italic, fontSize: 14 }],
                },
              ],
            },
          },
        };
      }),
      // 2 images per slide
      ...Array.from({ length: 2 }, (_, pi) => ({
        picture: {
          x: 50 + pi * 400,
          y: 600,
          width: 350,
          height: 250,
          data: LARGE_IMAGES[(si * 2 + pi) % LARGE_IMAGES.length],
          type: "jpg" as const,
        },
      })),
      // 3×3 table
      {
        table: {
          rows: Array.from({ length: 3 }, (_, rowIdx) => ({
            cells: Array.from({ length: 3 }, (_, colIdx) => ({
              text: `S${si + 1}R${rowIdx + 1}C${colIdx + 1}`,
              fill: rowIdx === 0 ? "4472C4" : undefined,
            })),
          })),
        },
      },
    ] as readonly SlideChild[],
  })),
});

const build30Slides10Images = (): PresentationOptions => ({
  slides: Array.from({ length: 30 }, (_, si) => ({
    children: [
      // 2 shapes per slide
      ...Array.from({ length: 2 }, (_, shi) => {
        const s = LARGE_SHAPES[(si * 2 + shi) % LARGE_SHAPES.length];
        return {
          shape: {
            x: 50 + shi * 400,
            y: 50,
            width: 350,
            height: 80,
            fill: s.fill,
            textBody: {
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [{ text: s.text, bold: s.bold, italic: s.italic, fontSize: 14 }],
                },
              ],
            },
          },
        };
      }),
      // 10 images per slide (500KB each)
      ...Array.from({ length: 10 }, (_, pi) => ({
        picture: {
          x: 50 + (pi % 5) * 180,
          y: 200 + Math.floor(pi / 5) * 140,
          width: 150,
          height: 110,
          data: LARGE_IMAGES[(si * 10 + pi) % LARGE_IMAGES.length],
          type: "jpg" as const,
        },
      })),
    ] as readonly SlideChild[],
  })),
});

describe("PPTX: Large Files — Create + toBuffer", () => {
  bench(
    "ours default sync — 30 slides × 20 shapes + toBufferSync",
    () => {
      generateSync(build30Slides20Shapes());
    },
    { iterations: 10 },
  );

  bench(
    "ours all-store sync — 30 slides × 20 shapes + toBuffer",
    () => {
      generateSync(build30Slides20Shapes(), "nodebuffer", { compression: { xml: 0 } });
    },
    { iterations: 10 },
  );

  bench(
    "ours default async — 30 slides × 20 shapes + toBuffer",
    async () => {
      await generate(build30Slides20Shapes());
    },
    { iterations: 10 },
  );

  bench(
    "ours all-store async — 30 slides × 20 shapes + toBuffer",
    async () => {
      await generate(build30Slides20Shapes(), "nodebuffer", { compression: { xml: 0 } });
    },
    { iterations: 10 },
  );

  bench(
    "pptxgenjs DEFLATE — 30 slides × 20 shapes + toBuffer",
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
    "pptxgenjs STORE — 30 slides × 20 shapes + toBuffer",
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
      await pptx.write({ outputType: "nodebuffer" });
    },
    { iterations: 10 },
  );

  bench(
    "ours default sync — 30 slides × 10 images + toBufferSync",
    () => {
      generateSync(build30Slides10Images());
    },
    { iterations: 10 },
  );

  bench(
    "ours all-store sync — 30 slides × 10 images + toBuffer",
    () => {
      generateSync(build30Slides10Images(), "nodebuffer", { compression: { xml: 0 } });
    },
    { iterations: 10 },
  );

  bench(
    "ours default async — 30 slides × 10 images + toBuffer",
    async () => {
      await generate(build30Slides10Images());
    },
    { iterations: 10 },
  );

  bench(
    "ours all-store async — 30 slides × 10 images + toBuffer",
    async () => {
      await generate(build30Slides10Images(), "nodebuffer", { compression: { xml: 0 } });
    },
    { iterations: 10 },
  );

  bench(
    "pptxgenjs DEFLATE — 30 slides × 10 images + toBuffer",
    async () => {
      const pptx = new PptxGenJS();
      for (let si = 0; si < 30; si++) {
        const slide = pptx.addSlide();
        for (let shi = 0; shi < 2; shi++) {
          const s = LARGE_SHAPES[(si * 2 + shi) % LARGE_SHAPES.length];
          slide.addText(
            [{ text: s.text, options: { bold: s.bold, italic: s.italic, fontSize: s.fontSize } }],
            {
              x: 0.5 + shi * 4,
              y: 0.5,
              w: 3.5,
              h: 0.8,
              fill: s.fill ? { color: s.fill } : undefined,
            },
          );
        }
        for (let pi = 0; pi < 10; pi++) {
          slide.addImage({
            data: LARGE_IMAGES_BASE64[(si * 10 + pi) % LARGE_IMAGES_BASE64.length],
            x: 0.5 + (pi % 5) * 2,
            y: 1.5 + Math.floor(pi / 5) * 2,
            w: 1.8,
            h: 1.5,
          });
        }
      }
      await pptx.write({ outputType: "nodebuffer", compression: true });
    },
    { iterations: 10 },
  );

  bench(
    "pptxgenjs STORE — 30 slides × 10 images + toBuffer",
    async () => {
      const pptx = new PptxGenJS();
      for (let si = 0; si < 30; si++) {
        const slide = pptx.addSlide();
        for (let shi = 0; shi < 2; shi++) {
          const s = LARGE_SHAPES[(si * 2 + shi) % LARGE_SHAPES.length];
          slide.addText(
            [{ text: s.text, options: { bold: s.bold, italic: s.italic, fontSize: s.fontSize } }],
            {
              x: 0.5 + shi * 4,
              y: 0.5,
              w: 3.5,
              h: 0.8,
              fill: s.fill ? { color: s.fill } : undefined,
            },
          );
        }
        for (let pi = 0; pi < 10; pi++) {
          slide.addImage({
            data: LARGE_IMAGES_BASE64[(si * 10 + pi) % LARGE_IMAGES_BASE64.length],
            x: 0.5 + (pi % 5) * 2,
            y: 1.5 + Math.floor(pi / 5) * 2,
            w: 1.8,
            h: 1.5,
          });
        }
      }
      await pptx.write({ outputType: "nodebuffer" });
    },
    { iterations: 10 },
  );

  bench(
    "ours default sync — 100x10 table + toBufferSync",
    () => {
      generateSync(build100x10Table());
    },
    { iterations: 10 },
  );

  bench(
    "ours all-store sync — 100x10 table + toBuffer",
    () => {
      generateSync(build100x10Table(), "nodebuffer", { compression: { xml: 0 } });
    },
    { iterations: 10 },
  );

  bench(
    "ours default async — 100x10 table + toBuffer",
    async () => {
      await generate(build100x10Table());
    },
    { iterations: 10 },
  );

  bench(
    "ours all-store async — 100x10 table + toBuffer",
    async () => {
      await generate(build100x10Table(), "nodebuffer", { compression: { xml: 0 } });
    },
    { iterations: 10 },
  );

  bench(
    "pptxgenjs DEFLATE — 100x10 table + toBuffer",
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
    "pptxgenjs STORE — 100x10 table + toBuffer",
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
      await pptx.write({ outputType: "nodebuffer" });
    },
    { iterations: 10 },
  );

  bench(
    "ours default sync — 50 slides full + toBufferSync",
    () => {
      generateSync(build50SlidesFull());
    },
    { iterations: 10 },
  );

  bench(
    "ours all-store sync — 50 slides full + toBuffer",
    () => {
      generateSync(build50SlidesFull(), "nodebuffer", { compression: { xml: 0 } });
    },
    { iterations: 10 },
  );

  bench(
    "ours default async — 50 slides full + toBuffer",
    async () => {
      await generate(build50SlidesFull());
    },
    { iterations: 10 },
  );

  bench(
    "ours all-store async — 50 slides full + toBuffer",
    async () => {
      await generate(build50SlidesFull(), "nodebuffer", { compression: { xml: 0 } });
    },
    { iterations: 10 },
  );

  bench(
    "pptxgenjs DEFLATE — 50 slides full + toBuffer",
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
        for (let pi = 0; pi < 2; pi++) {
          slide.addImage({
            data: LARGE_IMAGES_BASE64[(si * 2 + pi) % LARGE_IMAGES_BASE64.length],
            x: 0.5 + pi * 4,
            y: 5,
            w: 3,
            h: 2,
          });
        }
        slide.addTable(
          Array.from({ length: 3 }, (_, rowIdx) =>
            Array.from({ length: 3 }, (_, colIdx) => ({
              text: `S${si + 1}R${rowIdx + 1}C${colIdx + 1}`,
              options: rowIdx === 0 ? { fill: { color: "4472C4" }, color: "FFFFFF" } : undefined,
            })),
          ),
          { x: 7, y: 5, w: 2.5, h: 2 },
        );
      }
      await pptx.write({ outputType: "nodebuffer", compression: true });
    },
    { iterations: 10 },
  );

  bench(
    "pptxgenjs STORE — 50 slides full + toBuffer",
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
        for (let pi = 0; pi < 2; pi++) {
          slide.addImage({
            data: LARGE_IMAGES_BASE64[(si * 2 + pi) % LARGE_IMAGES_BASE64.length],
            x: 0.5 + pi * 4,
            y: 5,
            w: 3,
            h: 2,
          });
        }
        slide.addTable(
          Array.from({ length: 3 }, (_, rowIdx) =>
            Array.from({ length: 3 }, (_, colIdx) => ({
              text: `S${si + 1}R${rowIdx + 1}C${colIdx + 1}`,
              options: rowIdx === 0 ? { fill: { color: "4472C4" }, color: "FFFFFF" } : undefined,
            })),
          ),
          { x: 7, y: 5, w: 2.5, h: 2 },
        );
      }
      await pptx.write({ outputType: "nodebuffer" });
    },
    { iterations: 10 },
  );
});

// ── Large file ~100MB mixed benchmarks ──

// Mixed-size image pool: 1MB×10 + 2MB×10 + 3MB×10 + 5MB×8 = 100MB, 38 unique images
const MIXED_IMAGE_SIZES = [
  ...Array(10).fill(1024), // 1MB × 10
  ...Array(10).fill(2048), // 2MB × 10
  ...Array(10).fill(3072), // 3MB × 10
  ...Array(8).fill(5120), // 5MB × 8
];
const MIXED_IMAGES = MIXED_IMAGE_SIZES.map((sizeKB, i) => makeImage(i, sizeKB));
const MIXED_IMAGES_BASE64 = MIXED_IMAGES.map(
  (img) => `data:image/jpeg;base64,${Buffer.from(img).toString("base64")}`,
);

const buildMixed100MbPres = (): PresentationOptions => ({
  slides: Array.from({ length: 40 }, (_, si) => ({
    children: [
      // 2 shapes
      ...Array.from({ length: 2 }, (_, shi) => {
        const s = LARGE_SHAPES[(si * 2 + shi) % LARGE_SHAPES.length];
        return {
          shape: {
            x: 50 + shi * 380,
            y: 50,
            width: 340,
            height: 80,
            fill: s.fill,
            textBody: {
              children: [
                {
                  properties: { bullet: { type: "none" } },
                  children: [
                    {
                      text: s.text,
                      bold: s.bold,
                      italic: s.italic,
                      fontSize: s.fontSize,
                    },
                  ],
                },
              ],
            },
          },
        };
      }),
      // 2 images from mixed pool (cycled across 38 unique images)
      ...Array.from({ length: 2 }, (_, pi) => ({
        picture: {
          x: 50 + pi * 400,
          y: 200,
          width: 350,
          height: 280,
          data: MIXED_IMAGES[(si * 2 + pi) % MIXED_IMAGES.length],
          type: "jpg" as const,
        },
      })),
      // 3×3 table
      {
        table: {
          rows: Array.from({ length: 3 }, (_, rowIdx) => ({
            cells: Array.from({ length: 3 }, (_, colIdx) => ({
              text: `S${si + 1}R${rowIdx + 1}C${colIdx + 1}`,
              fill: rowIdx === 0 ? "4472C4" : undefined,
            })),
          })),
        },
      },
    ] as readonly SlideChild[],
  })),
});

describe("PPTX: Large File (~100MB) — Mixed + async vs sync", () => {
  bench(
    "ours default sync — mixed (40sl, 38img) + toBufferSync",
    () => {
      generateSync(buildMixed100MbPres());
    },
    { iterations: 3 },
  );

  bench(
    "ours all-store sync — mixed (40sl, 38img) + toBuffer",
    () => {
      generateSync(buildMixed100MbPres(), "nodebuffer", { compression: { xml: 0 } });
    },
    { iterations: 3 },
  );

  bench(
    "ours default async — mixed (40sl, 38img) + toBuffer",
    async () => {
      await generate(buildMixed100MbPres());
    },
    { iterations: 3 },
  );

  bench(
    "ours all-store async — mixed (40sl, 38img) + toBuffer",
    async () => {
      await generate(buildMixed100MbPres(), "nodebuffer", { compression: { xml: 0 } });
    },
    { iterations: 3 },
  );

  bench(
    "pptxgenjs DEFLATE — mixed (40sl, 38img) + toBuffer",
    async () => {
      const pptx = new PptxGenJS();
      for (let si = 0; si < 40; si++) {
        const slide = pptx.addSlide();
        // 2 shapes
        for (let shi = 0; shi < 2; shi++) {
          const s = LARGE_SHAPES[(si * 2 + shi) % LARGE_SHAPES.length];
          slide.addText(
            [{ text: s.text, options: { bold: s.bold, italic: s.italic, fontSize: s.fontSize } }],
            {
              x: 0.5 + shi * 3.8,
              y: 0.5,
              w: 3.4,
              h: 0.8,
              fill: s.fill ? { color: s.fill } : undefined,
            },
          );
        }
        // 2 images
        for (let pi = 0; pi < 2; pi++) {
          slide.addImage({
            data: MIXED_IMAGES_BASE64[(si * 2 + pi) % MIXED_IMAGES_BASE64.length],
            x: 0.5 + pi * 4,
            y: 2,
            w: 3.5,
            h: 2.8,
          });
        }
        // 3×3 table
        slide.addTable(
          Array.from({ length: 3 }, (_, rowIdx) =>
            Array.from({ length: 3 }, (_, colIdx) => ({
              text: `S${si + 1}R${rowIdx + 1}C${colIdx + 1}`,
              options: rowIdx === 0 ? { fill: { color: "4472C4" }, color: "FFFFFF" } : undefined,
            })),
          ),
          { x: 0.5, y: 5, w: 4, h: 2 },
        );
      }
      await pptx.write({ outputType: "nodebuffer", compression: true });
    },
    { iterations: 3 },
  );

  bench(
    "pptxgenjs STORE — mixed (40sl, 38img) + toBuffer",
    async () => {
      const pptx = new PptxGenJS();
      for (let si = 0; si < 40; si++) {
        const slide = pptx.addSlide();
        // 2 shapes
        for (let shi = 0; shi < 2; shi++) {
          const s = LARGE_SHAPES[(si * 2 + shi) % LARGE_SHAPES.length];
          slide.addText(
            [{ text: s.text, options: { bold: s.bold, italic: s.italic, fontSize: s.fontSize } }],
            {
              x: 0.5 + shi * 3.8,
              y: 0.5,
              w: 3.4,
              h: 0.8,
              fill: s.fill ? { color: s.fill } : undefined,
            },
          );
        }
        // 2 images
        for (let pi = 0; pi < 2; pi++) {
          slide.addImage({
            data: MIXED_IMAGES_BASE64[(si * 2 + pi) % MIXED_IMAGES_BASE64.length],
            x: 0.5 + pi * 4,
            y: 2,
            w: 3.5,
            h: 2.8,
          });
        }
        // 3×3 table
        slide.addTable(
          Array.from({ length: 3 }, (_, rowIdx) =>
            Array.from({ length: 3 }, (_, colIdx) => ({
              text: `S${si + 1}R${rowIdx + 1}C${colIdx + 1}`,
              options: rowIdx === 0 ? { fill: { color: "4472C4" }, color: "FFFFFF" } : undefined,
            })),
          ),
          { x: 0.5, y: 5, w: 4, h: 2 },
        );
      }
      await pptx.write({ outputType: "nodebuffer" });
    },
    { iterations: 3 },
  );
});
