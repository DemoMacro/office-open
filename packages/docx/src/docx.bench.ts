import {
  OoxmlMimeType,
  ZIP_STORED_LEVEL,
  zipAndConvert,
  zipSyncAndConvert,
} from "@office-open/core";
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

// STORE sync: compile + zip with STORE (no compression)
const toBufferStore = (doc: Document) => {
  const files = Packer.compile(doc);
  return zipSyncAndConvert(files, "nodebuffer", OoxmlMimeType.DOCX, ZIP_STORED_LEVEL);
};

// STORE async: compile + zip async with STORE
const toBufferStoreAsync = async (doc: Document) => {
  const files = Packer.compile(doc);
  return await zipAndConvert(files, "nodebuffer", OoxmlMimeType.DOCX, ZIP_STORED_LEVEL);
};

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
        ],
      },
    ],
  });

const buildStyledDoc = () =>
  new Document({
    sections: [
      {
        children: PARAGRAPH_CHILDREN.map(
          (p) =>
            new Paragraph({
              children: [new TextRun({ text: p.text, bold: p.bold, italics: p.italics })],
            }),
        ),
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
        ],
      },
    ],
  });

const buildStyledDocCompetitor = () =>
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

// ── Benchmarks ──

// Both libraries use DEFLATE compression for a fair comparison.
// Our Packer uses fflate zipSync() for maximum throughput; docx uses JSZip.
describe("DOCX: Create + toBuffer", () => {
  bench(
    "ours DEFLATE sync — simple + toBufferSync",
    () => {
      Packer.toBufferSync(buildSimpleDoc());
    },
    { iterations: 50 },
  );

  bench(
    "ours STORE sync — simple + toBufferStore",
    () => {
      toBufferStore(buildSimpleDoc());
    },
    { iterations: 50 },
  );

  bench(
    "ours DEFLATE async — simple + toBuffer",
    async () => {
      await Packer.toBuffer(buildSimpleDoc());
    },
    { iterations: 50 },
  );

  bench(
    "ours STORE async — simple + toBufferStoreAsync",
    async () => {
      await toBufferStoreAsync(buildSimpleDoc());
    },
    { iterations: 50 },
  );

  bench(
    "docx — simple + toBuffer",
    async () => {
      await PackerOrig.toBuffer(buildSimpleDocCompetitor());
    },
    { iterations: 50 },
  );

  bench(
    "ours DEFLATE sync — styled paragraphs (20) + toBufferSync",
    () => {
      Packer.toBufferSync(buildStyledDoc());
    },
    { iterations: 50 },
  );

  bench(
    "ours STORE sync — styled paragraphs (20) + toBufferStore",
    () => {
      toBufferStore(buildStyledDoc());
    },
    { iterations: 50 },
  );

  bench(
    "ours DEFLATE async — styled paragraphs (20) + toBuffer",
    async () => {
      await Packer.toBuffer(buildStyledDoc());
    },
    { iterations: 50 },
  );

  bench(
    "ours STORE async — styled paragraphs (20) + toBufferStoreAsync",
    async () => {
      await toBufferStoreAsync(buildStyledDoc());
    },
    { iterations: 50 },
  );

  bench(
    "docx — styled paragraphs (20) + toBuffer",
    async () => {
      await PackerOrig.toBuffer(buildStyledDocCompetitor());
    },
    { iterations: 50 },
  );

  bench(
    "ours DEFLATE sync — table (10x5) + toBufferSync",
    () => {
      Packer.toBufferSync(buildTableDoc());
    },
    { iterations: 50 },
  );

  bench(
    "ours STORE sync — table (10x5) + toBufferStore",
    () => {
      toBufferStore(buildTableDoc());
    },
    { iterations: 50 },
  );

  bench(
    "ours DEFLATE async — table (10x5) + toBuffer",
    async () => {
      await Packer.toBuffer(buildTableDoc());
    },
    { iterations: 50 },
  );

  bench(
    "ours STORE async — table (10x5) + toBufferStoreAsync",
    async () => {
      await toBufferStoreAsync(buildTableDoc());
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
    "ours DEFLATE sync — full featured + toBufferSync",
    () => {
      Packer.toBufferSync(buildFullFeaturedDoc());
    },
    { iterations: 50 },
  );

  bench(
    "ours STORE sync — full featured + toBufferStore",
    () => {
      toBufferStore(buildFullFeaturedDoc());
    },
    { iterations: 50 },
  );

  bench(
    "ours DEFLATE async — full featured + toBuffer",
    async () => {
      await Packer.toBuffer(buildFullFeaturedDoc());
    },
    { iterations: 50 },
  );

  bench(
    "ours STORE async — full featured + toBufferStoreAsync",
    async () => {
      await toBufferStoreAsync(buildFullFeaturedDoc());
    },
    { iterations: 50 },
  );

  bench(
    "docx — full featured + toBuffer",
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
      children: Array.from({ length: 100 }, (_, pi) => {
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
      children: Array.from({ length: 100 }, (_, pi) => {
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
    })),
  });

describe("DOCX: Large Files — Create + toBuffer", () => {
  bench(
    "ours DEFLATE sync — 2000 paragraphs + toBufferSync",
    () => {
      Packer.toBufferSync(buildLargeParagraphsDoc());
    },
    { iterations: 10 },
  );

  bench(
    "ours STORE sync — 2000 paragraphs + toBufferStore",
    () => {
      toBufferStore(buildLargeParagraphsDoc());
    },
    { iterations: 10 },
  );

  bench(
    "ours DEFLATE async — 2000 paragraphs + toBuffer",
    async () => {
      await Packer.toBuffer(buildLargeParagraphsDoc());
    },
    { iterations: 10 },
  );

  bench(
    "ours STORE async — 2000 paragraphs + toBufferStoreAsync",
    async () => {
      await toBufferStoreAsync(buildLargeParagraphsDoc());
    },
    { iterations: 10 },
  );

  bench(
    "docx — 2000 paragraphs + toBuffer",
    async () => {
      await PackerOrig.toBuffer(buildLargeParagraphsDocCompetitor());
    },
    { iterations: 10 },
  );

  bench(
    "ours DEFLATE sync — 200x10 table + toBufferSync",
    () => {
      Packer.toBufferSync(buildLargeTableDoc());
    },
    { iterations: 10 },
  );

  bench(
    "ours STORE sync — 200x10 table + toBufferStore",
    () => {
      toBufferStore(buildLargeTableDoc());
    },
    { iterations: 10 },
  );

  bench(
    "ours DEFLATE async — 200x10 table + toBuffer",
    async () => {
      await Packer.toBuffer(buildLargeTableDoc());
    },
    { iterations: 10 },
  );

  bench(
    "ours STORE async — 200x10 table + toBufferStoreAsync",
    async () => {
      await toBufferStoreAsync(buildLargeTableDoc());
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
    "ours DEFLATE sync — 20 sections × 100 paragraphs + toBufferSync",
    () => {
      Packer.toBufferSync(buildLargeSectionsDoc());
    },
    { iterations: 10 },
  );

  bench(
    "ours STORE sync — 20 sections × 100 paragraphs + toBufferStore",
    () => {
      toBufferStore(buildLargeSectionsDoc());
    },
    { iterations: 10 },
  );

  bench(
    "ours DEFLATE async — 20 sections × 100 paragraphs + toBuffer",
    async () => {
      await Packer.toBuffer(buildLargeSectionsDoc());
    },
    { iterations: 10 },
  );

  bench(
    "ours STORE async — 20 sections × 100 paragraphs + toBufferStoreAsync",
    async () => {
      await toBufferStoreAsync(buildLargeSectionsDoc());
    },
    { iterations: 10 },
  );

  bench(
    "docx — 20 sections × 100 paragraphs + toBuffer",
    async () => {
      await PackerOrig.toBuffer(buildLargeSectionsDocCompetitor());
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

const buildMixed100MbDoc = () =>
  new Document({
    sections: [
      {
        children: [
          // 2000 paragraphs (bold/italic alternating)
          ...LARGE_PARAGRAPHS.map(
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
          // 200 unique images (500KB each)
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
          // 100×10 table
          new Table({
            rows: Array.from(
              { length: 100 },
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
    "ours DEFLATE sync — mixed (2kp+200img+100x10) + toBufferSync",
    () => {
      Packer.toBufferSync(buildMixed100MbDoc());
    },
    { iterations: 3 },
  );

  bench(
    "ours STORE sync — mixed (2kp+200img+100x10) + toBufferStore",
    () => {
      toBufferStore(buildMixed100MbDoc());
    },
    { iterations: 3 },
  );

  bench(
    "ours DEFLATE async — mixed (2kp+200img+100x10) + toBuffer",
    async () => {
      await Packer.toBuffer(buildMixed100MbDoc());
    },
    { iterations: 3 },
  );

  bench(
    "ours STORE async — mixed (2kp+200img+100x10) + toBufferStoreAsync",
    async () => {
      await toBufferStoreAsync(buildMixed100MbDoc());
    },
    { iterations: 3 },
  );

  bench(
    "docx — mixed (2kp+200img+100x10) + toBuffer",
    async () => {
      const doc = new DocumentOrig({
        sections: [
          {
            children: [
              // 2000 paragraphs
              ...LARGE_PARAGRAPHS.map(
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
              // 200 unique images (500KB each)
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
              // 100×10 table
              new TableOrig({
                rows: Array.from(
                  { length: 100 },
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
