import {
  AlignmentType as AlignmentTypeOrig,
  Document as DocumentOrig,
  Footer as FooterOrig,
  Header as HeaderOrig,
  HeadingLevel as HeadingLevelOrig,
  ImageRun as ImageRunOrig,
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
  ImageRun,
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

// Bench modes:
//   "ours default"  = XML DEFLATE level 1 (SuperFast, MS Office), media STORE — no options passed.
//   "ours all-store" = all entries STORE — { compression: { xml: 0 } }.
//
// docx (JSZip): async (Packer.toBuffer). Hardcoded global DEFLATE for ALL entries,
// including images (redundant compression). No STORE option.

// ── Image generation ──

const makeImage = (seed: number, sizeKB: number): Uint8Array => {
  const size = sizeKB * 1024;
  const buf = new Uint8Array(size);
  for (let i = 0; i < size; i++) buf[i] = (i * 7 + seed * 13 + 37) & 0xff;
  return buf;
};

const SMALL_IMAGES = Array.from({ length: 3 }, (_, i) => makeImage(i, 200));
const LARGE_IMAGES = Array.from({ length: 20 }, (_, i) => makeImage(i, 500));

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

// ── Fixture helpers (ours) ──

const buildSimpleDoc = () =>
  new Document({
    sections: [
      {
        children: [
          new Paragraph({ children: [new TextRun("Hello World")] }),
          new Paragraph({ children: [new TextRun("Second paragraph")] }),
          new Paragraph({
            children: [
              new ImageRun({
                data: SMALL_IMAGES[0],
                transformation: { width: 400, height: 300 },
                type: "jpg",
              }),
            ],
          }),
        ],
      },
    ],
  });

const buildStyledDoc = () =>
  new Document({
    sections: [
      {
        children: [
          ...PARAGRAPH_CHILDREN.map(
            (p) =>
              new Paragraph({
                children: [new TextRun({ text: p.text, bold: p.bold, italics: p.italics })],
              }),
          ),
          new Paragraph({
            children: [
              new ImageRun({
                data: SMALL_IMAGES[1],
                transformation: { width: 400, height: 300 },
                type: "jpg",
              }),
            ],
          }),
        ],
      },
    ],
  });

const buildTableDoc = () =>
  new Document({
    sections: [
      {
        children: [
          new Table({
            rows: TABLE_ROWS.map(
              (row) =>
                new TableRow({
                  cells: row.cells.map(
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

const buildFullFeaturedDoc = () =>
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
          new Paragraph({
            children: [
              new ImageRun({
                data: SMALL_IMAGES[0],
                transformation: { width: 400, height: 300 },
                type: "jpg",
              }),
            ],
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
          new Paragraph({
            children: [
              new ImageRun({
                data: SMALL_IMAGES[2],
                transformation: { width: 400, height: 300 },
                type: "jpg",
              }),
            ],
          }),
          new Table({
            rows: TABLE_ROWS.map(
              (row) =>
                new TableRow({
                  cells: row.cells.map(
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

// ── Fixture helpers (competitor) ──

const buildSimpleDocCompetitor = () =>
  new DocumentOrig({
    sections: [
      {
        children: [
          new ParagraphOrig({ children: [new TextRunOrig("Hello World")] }),
          new ParagraphOrig({ children: [new TextRunOrig("Second paragraph")] }),
          new ParagraphOrig({
            children: [
              new ImageRunOrig({
                data: SMALL_IMAGES[0],
                transformation: { width: 400, height: 300 },
                type: "jpg",
              }),
            ],
          }),
        ],
      },
    ],
  });

const buildStyledDocCompetitor = () =>
  new DocumentOrig({
    sections: [
      {
        children: [
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
          new ParagraphOrig({
            children: [
              new ImageRunOrig({
                data: SMALL_IMAGES[1],
                transformation: { width: 400, height: 300 },
                type: "jpg",
              }),
            ],
          }),
        ],
      },
    ],
  });

const buildTableDocCompetitor = () =>
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

const buildFullFeaturedDocCompetitor = () =>
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
          new ParagraphOrig({
            children: [
              new ImageRunOrig({
                data: SMALL_IMAGES[0],
                transformation: { width: 400, height: 300 },
                type: "jpg",
              }),
            ],
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
          new ParagraphOrig({
            children: [
              new ImageRunOrig({
                data: SMALL_IMAGES[2],
                transformation: { width: 400, height: 300 },
                type: "jpg",
              }),
            ],
          }),
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

// ── Benchmarks ──

describe("DOCX: Create + toBuffer", () => {
  bench(
    "ours default sync — simple (2p + 1 img) + toBufferSync",
    () => {
      Packer.toBufferSync(buildSimpleDoc());
    },
    { iterations: 50 },
  );

  bench(
    "ours all-store sync — simple (2p + 1 img) + toBufferStore",
    () => {
      Packer.toBufferSync(buildSimpleDoc(), { compression: { xml: 0 } });
    },
    { iterations: 50 },
  );

  bench(
    "ours default async — simple (2p + 1 img) + toBuffer",
    async () => {
      await Packer.toBuffer(buildSimpleDoc());
    },
    { iterations: 50 },
  );

  bench(
    "ours all-store async — simple (2p + 1 img) + toBufferStoreAsync",
    async () => {
      await Packer.toBuffer(buildSimpleDoc(), { compression: { xml: 0 } });
    },
    { iterations: 50 },
  );

  bench(
    "docx — simple (2p + 1 img) + toBuffer",
    async () => {
      await PackerOrig.toBuffer(buildSimpleDocCompetitor());
    },
    { iterations: 50 },
  );

  bench(
    "ours default sync — styled paragraphs (20) + 1 img + toBufferSync",
    () => {
      Packer.toBufferSync(buildStyledDoc());
    },
    { iterations: 50 },
  );

  bench(
    "ours all-store sync — styled paragraphs (20) + 1 img + toBufferStore",
    () => {
      Packer.toBufferSync(buildStyledDoc(), { compression: { xml: 0 } });
    },
    { iterations: 50 },
  );

  bench(
    "ours default async — styled paragraphs (20) + 1 img + toBuffer",
    async () => {
      await Packer.toBuffer(buildStyledDoc());
    },
    { iterations: 50 },
  );

  bench(
    "ours all-store async — styled paragraphs (20) + 1 img + toBufferStoreAsync",
    async () => {
      await Packer.toBuffer(buildStyledDoc(), { compression: { xml: 0 } });
    },
    { iterations: 50 },
  );

  bench(
    "docx — styled paragraphs (20) + 1 img + toBuffer",
    async () => {
      await PackerOrig.toBuffer(buildStyledDocCompetitor());
    },
    { iterations: 50 },
  );

  bench(
    "ours default sync — table (10x5) + toBufferSync",
    () => {
      Packer.toBufferSync(buildTableDoc());
    },
    { iterations: 50 },
  );

  bench(
    "ours all-store sync — table (10x5) + toBufferStore",
    () => {
      Packer.toBufferSync(buildTableDoc(), { compression: { xml: 0 } });
    },
    { iterations: 50 },
  );

  bench(
    "ours default async — table (10x5) + toBuffer",
    async () => {
      await Packer.toBuffer(buildTableDoc());
    },
    { iterations: 50 },
  );

  bench(
    "ours all-store async — table (10x5) + toBufferStoreAsync",
    async () => {
      await Packer.toBuffer(buildTableDoc(), { compression: { xml: 0 } });
    },
    { iterations: 50 },
  );

  bench(
    "docx — table (10x5) + toBuffer",
    async () => {
      await PackerOrig.toBuffer(buildTableDocCompetitor());
    },
    { iterations: 50 },
  );

  bench(
    "ours default sync — full featured + 2 imgs + toBufferSync",
    () => {
      Packer.toBufferSync(buildFullFeaturedDoc());
    },
    { iterations: 50 },
  );

  bench(
    "ours all-store sync — full featured + 2 imgs + toBufferStore",
    () => {
      Packer.toBufferSync(buildFullFeaturedDoc(), { compression: { xml: 0 } });
    },
    { iterations: 50 },
  );

  bench(
    "ours default async — full featured + 2 imgs + toBuffer",
    async () => {
      await Packer.toBuffer(buildFullFeaturedDoc());
    },
    { iterations: 50 },
  );

  bench(
    "ours all-store async — full featured + 2 imgs + toBufferStoreAsync",
    async () => {
      await Packer.toBuffer(buildFullFeaturedDoc(), { compression: { xml: 0 } });
    },
    { iterations: 50 },
  );

  bench(
    "docx — full featured + 2 imgs + toBuffer",
    async () => {
      await PackerOrig.toBuffer(buildFullFeaturedDocCompetitor());
    },
    { iterations: 50 },
  );
});

// ── Large file benchmarks ──

const LARGE_PARAGRAPHS = Array.from({ length: 2000 }, (_, i) => ({
  text: `This is paragraph ${i + 1} with sample content to simulate real document generation. Each paragraph contains enough text to represent a realistic scenario for stress testing large file creation.`,
  bold: i % 3 === 0,
  italics: i % 5 === 0,
  underline: i % 7 === 0,
}));

const LARGE_TABLE_ROWS = Array.from({ length: 200 }, (_, rowIdx) => ({
  cells: Array.from({ length: 10 }, (_, colIdx) => ({
    text: `R${rowIdx + 1}C${colIdx + 1} data content`,
    width: { size: 1000, type: WidthType.PERCENTAGE },
  })),
}));

const buildLargeParagraphsDoc = () =>
  new Document({
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
        ).flatMap((para, pi) =>
          pi > 0 && pi % 100 === 0
            ? [
                para,
                new Paragraph({
                  children: [
                    new ImageRun({
                      data: LARGE_IMAGES[(pi / 100 - 1) % LARGE_IMAGES.length],
                      transformation: { width: 400, height: 300 },
                      type: "jpg",
                    }),
                  ],
                }),
              ]
            : [para],
        ),
      },
    ],
  });

const buildLargeParagraphsDocCompetitor = () =>
  new DocumentOrig({
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
        ).flatMap((para, pi) =>
          pi > 0 && pi % 100 === 0
            ? [
                para,
                new ParagraphOrig({
                  children: [
                    new ImageRunOrig({
                      data: LARGE_IMAGES[(pi / 100 - 1) % LARGE_IMAGES.length],
                      transformation: { width: 400, height: 300 },
                      type: "jpg",
                    }),
                  ],
                }),
              ]
            : [para],
        ),
      },
    ],
  });

const buildLargeTableDoc = () =>
  new Document({
    sections: [
      {
        children: [
          new Table({
            rows: LARGE_TABLE_ROWS.map(
              (row) =>
                new TableRow({
                  cells: row.cells.map(
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

const buildLargeTableDocCompetitor = () =>
  new DocumentOrig({
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

const buildLargeSectionsDoc = () =>
  new Document({
    sections: Array.from({ length: 20 }, (_, si) => ({
      properties: { page: { margin: { top: 1440, bottom: 1440 } } },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [
                new TextRun(`Chapter ${si + 1}`),
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
              children: [new TextRun("Footer text")],
            }),
          ],
        }),
      },
      children: [
        ...Array.from({ length: 100 }, (_, pi) => {
          return new Paragraph({
            heading:
              pi === 0 ? HeadingLevel.HEADING_1 : pi === 1 ? HeadingLevel.HEADING_2 : undefined,
            children: [
              new TextRun({
                text:
                  pi === 0
                    ? `Chapter ${si + 1} Title`
                    : pi === 1
                      ? `Section ${si + 1}.${1} Subtitle`
                      : `Chapter ${si + 1} paragraph ${pi} body content for realistic document simulation with enough text.`,
                bold: pi <= 1,
              }),
            ],
          });
        }),
        new Paragraph({
          children: [
            new ImageRun({
              data: LARGE_IMAGES[(si * 2) % LARGE_IMAGES.length],
              transformation: { width: 400, height: 300 },
              type: "jpg",
            }),
          ],
        }),
        new Paragraph({
          children: [
            new ImageRun({
              data: LARGE_IMAGES[(si * 2 + 1) % LARGE_IMAGES.length],
              transformation: { width: 400, height: 300 },
              type: "jpg",
            }),
          ],
        }),
      ],
    })),
  });

const buildLargeSectionsDocCompetitor = () =>
  new DocumentOrig({
    sections: Array.from({ length: 20 }, (_, si) => ({
      properties: { page: { margin: { top: 1440, bottom: 1440 } } },
      headers: {
        default: new HeaderOrig({
          children: [
            new ParagraphOrig({
              alignment: AlignmentTypeOrig.RIGHT,
              children: [
                new TextRunOrig(`Chapter ${si + 1}`),
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
              children: [new TextRunOrig("Footer text")],
            }),
          ],
        }),
      },
      children: [
        ...Array.from({ length: 100 }, (_, pi) => {
          return new ParagraphOrig({
            heading:
              pi === 0
                ? HeadingLevelOrig.HEADING_1
                : pi === 1
                  ? HeadingLevelOrig.HEADING_2
                  : undefined,
            children: [
              new TextRunOrig({
                text:
                  pi === 0
                    ? `Chapter ${si + 1} Title`
                    : pi === 1
                      ? `Section ${si + 1}.${1} Subtitle`
                      : `Chapter ${si + 1} paragraph ${pi} body content for realistic document simulation with enough text.`,
                bold: pi <= 1,
              }),
            ],
          });
        }),
        new ParagraphOrig({
          children: [
            new ImageRunOrig({
              data: LARGE_IMAGES[(si * 2) % LARGE_IMAGES.length],
              transformation: { width: 400, height: 300 },
              type: "jpg",
            }),
          ],
        }),
        new ParagraphOrig({
          children: [
            new ImageRunOrig({
              data: LARGE_IMAGES[(si * 2 + 1) % LARGE_IMAGES.length],
              transformation: { width: 400, height: 300 },
              type: "jpg",
            }),
          ],
        }),
      ],
    })),
  });

describe("DOCX: Large Files — Create + toBuffer", () => {
  bench(
    "ours default sync — 2000p + 20 img + toBufferSync",
    () => {
      Packer.toBufferSync(buildLargeParagraphsDoc());
    },
    { iterations: 10 },
  );

  bench(
    "ours all-store sync — 2000p + 20 img + toBufferStore",
    () => {
      Packer.toBufferSync(buildLargeParagraphsDoc(), { compression: { xml: 0 } });
    },
    { iterations: 10 },
  );

  bench(
    "ours default async — 2000p + 20 img + toBuffer",
    async () => {
      await Packer.toBuffer(buildLargeParagraphsDoc());
    },
    { iterations: 10 },
  );

  bench(
    "ours all-store async — 2000p + 20 img + toBufferStoreAsync",
    async () => {
      await Packer.toBuffer(buildLargeParagraphsDoc(), { compression: { xml: 0 } });
    },
    { iterations: 10 },
  );

  bench(
    "docx — 2000p + 20 img + toBuffer",
    async () => {
      await PackerOrig.toBuffer(buildLargeParagraphsDocCompetitor());
    },
    { iterations: 10 },
  );

  bench(
    "ours default sync — 200x10 table + toBufferSync",
    () => {
      Packer.toBufferSync(buildLargeTableDoc());
    },
    { iterations: 10 },
  );

  bench(
    "ours all-store sync — 200x10 table + toBufferStore",
    () => {
      Packer.toBufferSync(buildLargeTableDoc(), { compression: { xml: 0 } });
    },
    { iterations: 10 },
  );

  bench(
    "ours default async — 200x10 table + toBuffer",
    async () => {
      await Packer.toBuffer(buildLargeTableDoc());
    },
    { iterations: 10 },
  );

  bench(
    "ours all-store async — 200x10 table + toBufferStoreAsync",
    async () => {
      await Packer.toBuffer(buildLargeTableDoc(), { compression: { xml: 0 } });
    },
    { iterations: 10 },
  );

  bench(
    "docx — 200x10 table + toBuffer",
    async () => {
      await PackerOrig.toBuffer(buildLargeTableDocCompetitor());
    },
    { iterations: 10 },
  );

  bench(
    "ours default sync — 20 sec × 100p + 40 img + toBufferSync",
    () => {
      Packer.toBufferSync(buildLargeSectionsDoc());
    },
    { iterations: 10 },
  );

  bench(
    "ours all-store sync — 20 sec × 100p + 40 img + toBufferStore",
    () => {
      Packer.toBufferSync(buildLargeSectionsDoc(), { compression: { xml: 0 } });
    },
    { iterations: 10 },
  );

  bench(
    "ours default async — 20 sec × 100p + 40 img + toBuffer",
    async () => {
      await Packer.toBuffer(buildLargeSectionsDoc());
    },
    { iterations: 10 },
  );

  bench(
    "ours all-store async — 20 sec × 100p + 40 img + toBufferStoreAsync",
    async () => {
      await Packer.toBuffer(buildLargeSectionsDoc(), { compression: { xml: 0 } });
    },
    { iterations: 10 },
  );

  bench(
    "docx — 20 sec × 100p + 40 img + toBuffer",
    async () => {
      await PackerOrig.toBuffer(buildLargeSectionsDocCompetitor());
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

const buildMixed100MbDoc = () =>
  new Document({
    sections: [
      {
        children: [
          // 500 styled paragraphs
          ...LARGE_PARAGRAPHS.slice(0, 500).map(
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
          // 38 mixed-size images
          ...MIXED_IMAGES.map(
            (img) =>
              new Paragraph({
                children: [
                  new ImageRun({
                    data: img,
                    transformation: { width: 400, height: 300 },
                    type: "jpg",
                  }),
                ],
              }),
          ),
          // 50×10 table
          new Table({
            rows: Array.from(
              { length: 50 },
              (_, rowIdx) =>
                new TableRow({
                  cells: Array.from(
                    { length: 10 },
                    (_, colIdx) =>
                      new TableCell({
                        width: { size: 1000, type: WidthType.PERCENTAGE },
                        children: [
                          new Paragraph({
                            children: [new TextRun(`R${rowIdx + 1}C${colIdx + 1} data content`)],
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

describe("DOCX: Large File (~100MB) — Mixed + async vs sync", () => {
  bench(
    "ours default sync — mixed (500p, 38img, 50x10) + toBufferSync",
    () => {
      Packer.toBufferSync(buildMixed100MbDoc());
    },
    { iterations: 3 },
  );

  bench(
    "ours all-store sync — mixed (500p, 38img, 50x10) + toBufferStore",
    () => {
      Packer.toBufferSync(buildMixed100MbDoc(), { compression: { xml: 0 } });
    },
    { iterations: 3 },
  );

  bench(
    "ours default async — mixed (500p, 38img, 50x10) + toBuffer",
    async () => {
      await Packer.toBuffer(buildMixed100MbDoc());
    },
    { iterations: 3 },
  );

  bench(
    "ours all-store async — mixed (500p, 38img, 50x10) + toBufferStoreAsync",
    async () => {
      await Packer.toBuffer(buildMixed100MbDoc(), { compression: { xml: 0 } });
    },
    { iterations: 3 },
  );

  bench(
    "docx — mixed (500p, 38img, 50x10) + toBuffer",
    async () => {
      const doc = new DocumentOrig({
        sections: [
          {
            children: [
              // 500 styled paragraphs
              ...LARGE_PARAGRAPHS.slice(0, 500).map(
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
              // 38 mixed-size images
              ...MIXED_IMAGES.map(
                (img) =>
                  new ParagraphOrig({
                    children: [
                      new ImageRunOrig({
                        data: img,
                        transformation: { width: 400, height: 300 },
                        type: "jpg",
                      }),
                    ],
                  }),
              ),
              // 50×10 table
              new TableOrig({
                rows: Array.from(
                  { length: 50 },
                  (_, rowIdx) =>
                    new TableRowOrig({
                      children: Array.from(
                        { length: 10 },
                        (_, colIdx) =>
                          new TableCellOrig({
                            width: { size: 1000, type: WidthType.PERCENTAGE },
                            children: [
                              new ParagraphOrig({
                                children: [
                                  new TextRunOrig(`R${rowIdx + 1}C${colIdx + 1} data content`),
                                ],
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
    },
    { iterations: 3 },
  );
});
