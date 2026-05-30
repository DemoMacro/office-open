import { writeXlsx as hucreWriteXlsx } from "hucre";
import type { WriteSheet as HucreWriteSheet } from "hucre";
import { bench, describe } from "vite-plus/test";

import { Workbook, Packer } from "./index";

// ── Shared fixture data ──

const SIMPLE_ROWS = [
  ["Name", "Age", "City"],
  ["Alice", 30, "New York"],
  ["Bob", 25, "London"],
];

const STYLED_ROWS = Array.from({ length: 20 }, (_, i) => ({
  name: `Person ${i + 1}`,
  score: Math.round(Math.random() * 100),
  active: i % 2 === 0,
}));

const TABLE_ROWS = Array.from({ length: 10 }, (_, rowIdx) =>
  Array.from({ length: 5 }, (_, colIdx) => `R${rowIdx + 1}C${colIdx + 1}`),
);

// ── Benchmarks ──

describe("XLSX: Create + toBuffer (sync)", () => {
  bench("ours sync — simple + toBufferSync", () => {
    const wb = new Workbook({
      worksheets: [
        {
          children: SIMPLE_ROWS.map((row) => ({
            cells: row.map((v) => ({ value: v })),
          })),
        },
      ],
    });
    Packer.toBufferSync(wb);
  });

  bench("hucre — simple + toBuffer", async () => {
    await hucreWriteXlsx({
      sheets: [{ name: "Sheet1", rows: SIMPLE_ROWS }],
    });
  });

  bench("ours sync — styled rows (20) + toBufferSync", () => {
    const wb = new Workbook({
      worksheets: [
        {
          children: STYLED_ROWS.map((r) => ({
            cells: [{ value: r.name }, { value: r.score }, { value: r.active }],
          })),
        },
      ],
    });
    Packer.toBufferSync(wb);
  });

  bench("hucre — styled rows (20) + toBuffer", async () => {
    await hucreWriteXlsx({
      sheets: [
        {
          name: "Sheet1",
          rows: STYLED_ROWS.map((r) => [r.name, r.score, r.active]),
        },
      ],
    });
  });

  bench("ours sync — table (10x5) + toBufferSync", () => {
    const wb = new Workbook({
      worksheets: [
        {
          children: TABLE_ROWS.map((row) => ({
            cells: row.map((v) => ({ value: v })),
          })),
        },
      ],
    });
    Packer.toBufferSync(wb);
  });

  bench("hucre — table (10x5) + toBuffer", async () => {
    await hucreWriteXlsx({
      sheets: [{ name: "Sheet1", rows: TABLE_ROWS }],
    });
  });
});

describe("XLSX: Create + toBuffer (async)", () => {
  bench("ours async — simple + toBuffer", async () => {
    const wb = new Workbook({
      worksheets: [
        {
          children: SIMPLE_ROWS.map((row) => ({
            cells: row.map((v) => ({ value: v })),
          })),
        },
      ],
    });
    await Packer.toBuffer(wb);
  });

  bench("ours async — styled rows (20) + toBuffer", async () => {
    const wb = new Workbook({
      worksheets: [
        {
          children: STYLED_ROWS.map((r) => ({
            cells: [{ value: r.name }, { value: r.score }, { value: r.active }],
          })),
        },
      ],
    });
    await Packer.toBuffer(wb);
  });

  bench("ours async — table (10x5) + toBuffer", async () => {
    const wb = new Workbook({
      worksheets: [
        {
          children: TABLE_ROWS.map((row) => ({
            cells: row.map((v) => ({ value: v })),
          })),
        },
      ],
    });
    await Packer.toBuffer(wb);
  });
});

// ── Large file benchmarks ──

const LARGE_ROWS = Array.from({ length: 2000 }, (_, i) => [
  `Row ${i + 1}`,
  Math.round(Math.random() * 1000),
  i % 2 === 0,
  `Data ${i + 1} content for realistic spreadsheet simulation`,
]);

const LARGE_TABLE_ROWS = Array.from({ length: 200 }, (_, rowIdx) =>
  Array.from({ length: 10 }, (_, colIdx) => `R${rowIdx + 1}C${colIdx + 1} data content`),
);

describe("XLSX: Large Files — Create + toBuffer", () => {
  bench("ours sync — 2000 rows + toBufferSync", () => {
    const wb = new Workbook({
      worksheets: [
        {
          children: LARGE_ROWS.map((row) => ({
            cells: row.map((v) => ({ value: v })),
          })),
        },
      ],
    });
    Packer.toBufferSync(wb);
  });

  bench("ours async — 2000 rows + toBuffer", async () => {
    const wb = new Workbook({
      worksheets: [
        {
          children: LARGE_ROWS.map((row) => ({
            cells: row.map((v) => ({ value: v })),
          })),
        },
      ],
    });
    await Packer.toBuffer(wb);
  });

  bench("hucre — 2000 rows + toBuffer", async () => {
    await hucreWriteXlsx({
      sheets: [{ name: "Sheet1", rows: LARGE_ROWS }],
    });
  });

  bench("ours sync — 200x10 table + toBufferSync", () => {
    const wb = new Workbook({
      worksheets: [
        {
          children: LARGE_TABLE_ROWS.map((row) => ({
            cells: row.map((v) => ({ value: v })),
          })),
        },
      ],
    });
    Packer.toBufferSync(wb);
  });

  bench("ours async — 200x10 table + toBuffer", async () => {
    const wb = new Workbook({
      worksheets: [
        {
          children: LARGE_TABLE_ROWS.map((row) => ({
            cells: row.map((v) => ({ value: v })),
          })),
        },
      ],
    });
    await Packer.toBuffer(wb);
  });

  bench("hucre — 200x10 table + toBuffer", async () => {
    await hucreWriteXlsx({
      sheets: [{ name: "Sheet1", rows: LARGE_TABLE_ROWS }],
    });
  });

  bench("ours sync — 20 sheets × 100 rows + toBufferSync", () => {
    const wb = new Workbook({
      worksheets: Array.from({ length: 20 }, (_, si) => ({
        name: `Sheet${si + 1}`,
        children: Array.from({ length: 100 }, (_, ri) => ({
          cells: [
            { value: `S${si + 1}R${ri + 1}` },
            { value: ri * 10 + si },
            { value: `Data for sheet ${si + 1} row ${ri + 1}` },
          ],
        })),
      })),
    });
    Packer.toBufferSync(wb);
  });

  bench("ours async — 20 sheets × 100 rows + toBuffer", async () => {
    const wb = new Workbook({
      worksheets: Array.from({ length: 20 }, (_, si) => ({
        name: `Sheet${si + 1}`,
        children: Array.from({ length: 100 }, (_, ri) => ({
          cells: [
            { value: `S${si + 1}R${ri + 1}` },
            { value: ri * 10 + si },
            { value: `Data for sheet ${si + 1} row ${ri + 1}` },
          ],
        })),
      })),
    });
    await Packer.toBuffer(wb);
  });

  bench("hucre — 20 sheets × 100 rows + toBuffer", async () => {
    const sheets: HucreWriteSheet[] = Array.from({ length: 20 }, (_, si) => ({
      name: `Sheet${si + 1}`,
      rows: Array.from({ length: 100 }, (_, ri) => [
        `S${si + 1}R${ri + 1}`,
        ri * 10 + si,
        `Data for sheet ${si + 1} row ${ri + 1}`,
      ]),
    }));
    await hucreWriteXlsx({ sheets });
  });
});

// ── Large file ~100MB mixed benchmarks ──

const makeImage = (seed: number): Uint8Array => {
  const size = 500 * 1024;
  const buf = new Uint8Array(size);
  for (let i = 0; i < size; i++) buf[i] = (i * 7 + seed * 13 + 37) & 0xff;
  return buf;
};

const MIXED_IMAGES = Array.from({ length: 200 }, (_, i) => makeImage(i));

const buildMixed100MbWb = () =>
  new Workbook({
    worksheets: [
      {
        children: LARGE_ROWS.map((row) => ({
          cells: row.map((v) => ({ value: v })),
        })),
        images: MIXED_IMAGES.map((img, i) => ({
          data: img,
          type: "png" as const,
          col: (i % 10) + 5,
          row: Math.floor(i / 10) + 1,
        })),
      },
    ],
  });

describe("XLSX: Large File (~100MB) — Mixed + async vs sync", () => {
  bench("ours sync — mixed (2kp+200img) + toBufferSync", () => {
    Packer.toBufferSync(buildMixed100MbWb());
  });

  bench("ours async — mixed (2kp+200img) + toBuffer", async () => {
    await Packer.toBuffer(buildMixed100MbWb());
  });

  bench("hucre — mixed (2kp+200img) + toBuffer", async () => {
    const sheets: HucreWriteSheet[] = [
      {
        name: "Sheet1",
        rows: LARGE_ROWS,
        images: MIXED_IMAGES.map((img, i) => ({
          data: img,
          type: "png" as const,
          anchor: {
            from: { row: Math.floor(i / 10), col: (i % 10) + 5 },
            to: { row: Math.floor(i / 10) + 1, col: (i % 10) + 6 },
          },
        })),
      },
    ];
    await hucreWriteXlsx({ sheets });
  });
});
