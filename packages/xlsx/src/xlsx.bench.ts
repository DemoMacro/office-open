import {
  OoxmlMimeType,
  ZIP_STORED_LEVEL,
  zipAndConvert,
  zipSyncAndConvert,
} from "@office-open/core";
import { writeXlsx as hucreWriteXlsx } from "hucre";
import type { WriteSheet as HucreWriteSheet } from "hucre";
import { bench, describe } from "vite-plus/test";

import { Workbook, Packer } from "./index";

// STORE sync: compile + zip with STORE (no compression)
const toBufferStore = (wb: Workbook) => {
  const files = Packer.compile(wb);
  return zipSyncAndConvert(files, "nodebuffer", OoxmlMimeType.XLSX, ZIP_STORED_LEVEL);
};

// STORE async: compile + zip async with STORE
const toBufferStoreAsync = async (wb: Workbook) => {
  const files = Packer.compile(wb);
  return await zipAndConvert(files, "nodebuffer", OoxmlMimeType.XLSX, ZIP_STORED_LEVEL);
};

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

// ── Fixture helpers (ours) ──

const buildSimpleWb = () =>
  new Workbook({
    worksheets: [
      {
        children: SIMPLE_ROWS.map((row) => ({
          cells: row.map((v) => ({ value: v })),
        })),
      },
    ],
  });

const buildStyledWb = () =>
  new Workbook({
    worksheets: [
      {
        children: STYLED_ROWS.map((r) => ({
          cells: [{ value: r.name }, { value: r.score }, { value: r.active }],
        })),
      },
    ],
  });

const buildTableWb = () =>
  new Workbook({
    worksheets: [
      {
        children: TABLE_ROWS.map((row) => ({
          cells: row.map((v) => ({ value: v })),
        })),
      },
    ],
  });

// ── Benchmarks ──

describe("XLSX: Create + toBuffer", () => {
  bench(
    "ours DEFLATE sync — simple + toBufferSync",
    () => {
      Packer.toBufferSync(buildSimpleWb());
    },
    { iterations: 50 },
  );

  bench(
    "ours STORE sync — simple + toBufferStore",
    () => {
      toBufferStore(buildSimpleWb());
    },
    { iterations: 50 },
  );

  bench(
    "ours DEFLATE async — simple + toBuffer",
    async () => {
      await Packer.toBuffer(buildSimpleWb());
    },
    { iterations: 50 },
  );

  bench(
    "ours STORE async — simple + toBufferStoreAsync",
    async () => {
      await toBufferStoreAsync(buildSimpleWb());
    },
    { iterations: 50 },
  );

  bench(
    "hucre — simple + toBuffer",
    async () => {
      await hucreWriteXlsx({
        sheets: [{ name: "Sheet1", rows: SIMPLE_ROWS }],
      });
    },
    { iterations: 50 },
  );

  bench(
    "ours DEFLATE sync — styled rows (20) + toBufferSync",
    () => {
      Packer.toBufferSync(buildStyledWb());
    },
    { iterations: 50 },
  );

  bench(
    "ours STORE sync — styled rows (20) + toBufferStore",
    () => {
      toBufferStore(buildStyledWb());
    },
    { iterations: 50 },
  );

  bench(
    "ours DEFLATE async — styled rows (20) + toBuffer",
    async () => {
      await Packer.toBuffer(buildStyledWb());
    },
    { iterations: 50 },
  );

  bench(
    "ours STORE async — styled rows (20) + toBufferStoreAsync",
    async () => {
      await toBufferStoreAsync(buildStyledWb());
    },
    { iterations: 50 },
  );

  bench(
    "hucre — styled rows (20) + toBuffer",
    async () => {
      await hucreWriteXlsx({
        sheets: [
          {
            name: "Sheet1",
            rows: STYLED_ROWS.map((r) => [r.name, r.score, r.active]),
          },
        ],
      });
    },
    { iterations: 50 },
  );

  bench(
    "ours DEFLATE sync — table (10x5) + toBufferSync",
    () => {
      Packer.toBufferSync(buildTableWb());
    },
    { iterations: 50 },
  );

  bench(
    "ours STORE sync — table (10x5) + toBufferStore",
    () => {
      toBufferStore(buildTableWb());
    },
    { iterations: 50 },
  );

  bench(
    "ours DEFLATE async — table (10x5) + toBuffer",
    async () => {
      await Packer.toBuffer(buildTableWb());
    },
    { iterations: 50 },
  );

  bench(
    "ours STORE async — table (10x5) + toBufferStoreAsync",
    async () => {
      await toBufferStoreAsync(buildTableWb());
    },
    { iterations: 50 },
  );

  bench(
    "hucre — table (10x5) + toBuffer",
    async () => {
      await hucreWriteXlsx({
        sheets: [{ name: "Sheet1", rows: TABLE_ROWS }],
      });
    },
    { iterations: 50 },
  );
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

const buildLargeRowsWb = () =>
  new Workbook({
    worksheets: [
      {
        children: LARGE_ROWS.map((row) => ({
          cells: row.map((v) => ({ value: v })),
        })),
      },
    ],
  });

const buildLargeTableWb = () =>
  new Workbook({
    worksheets: [
      {
        children: LARGE_TABLE_ROWS.map((row) => ({
          cells: row.map((v) => ({ value: v })),
        })),
      },
    ],
  });

const buildLargeSheetsWb = () =>
  new Workbook({
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

describe("XLSX: Large Files — Create + toBuffer", () => {
  bench(
    "ours DEFLATE sync — 2000 rows + toBufferSync",
    () => {
      Packer.toBufferSync(buildLargeRowsWb());
    },
    { iterations: 10 },
  );

  bench(
    "ours STORE sync — 2000 rows + toBufferStore",
    () => {
      toBufferStore(buildLargeRowsWb());
    },
    { iterations: 10 },
  );

  bench(
    "ours DEFLATE async — 2000 rows + toBuffer",
    async () => {
      await Packer.toBuffer(buildLargeRowsWb());
    },
    { iterations: 10 },
  );

  bench(
    "ours STORE async — 2000 rows + toBufferStoreAsync",
    async () => {
      await toBufferStoreAsync(buildLargeRowsWb());
    },
    { iterations: 10 },
  );

  bench(
    "hucre — 2000 rows + toBuffer",
    async () => {
      await hucreWriteXlsx({
        sheets: [{ name: "Sheet1", rows: LARGE_ROWS }],
      });
    },
    { iterations: 10 },
  );

  bench(
    "ours DEFLATE sync — 200x10 table + toBufferSync",
    () => {
      Packer.toBufferSync(buildLargeTableWb());
    },
    { iterations: 10 },
  );

  bench(
    "ours STORE sync — 200x10 table + toBufferStore",
    () => {
      toBufferStore(buildLargeTableWb());
    },
    { iterations: 10 },
  );

  bench(
    "ours DEFLATE async — 200x10 table + toBuffer",
    async () => {
      await Packer.toBuffer(buildLargeTableWb());
    },
    { iterations: 10 },
  );

  bench(
    "ours STORE async — 200x10 table + toBufferStoreAsync",
    async () => {
      await toBufferStoreAsync(buildLargeTableWb());
    },
    { iterations: 10 },
  );

  bench(
    "hucre — 200x10 table + toBuffer",
    async () => {
      await hucreWriteXlsx({
        sheets: [{ name: "Sheet1", rows: LARGE_TABLE_ROWS }],
      });
    },
    { iterations: 10 },
  );

  bench(
    "ours DEFLATE sync — 20 sheets × 100 rows + toBufferSync",
    () => {
      Packer.toBufferSync(buildLargeSheetsWb());
    },
    { iterations: 10 },
  );

  bench(
    "ours STORE sync — 20 sheets × 100 rows + toBufferStore",
    () => {
      toBufferStore(buildLargeSheetsWb());
    },
    { iterations: 10 },
  );

  bench(
    "ours DEFLATE async — 20 sheets × 100 rows + toBuffer",
    async () => {
      await Packer.toBuffer(buildLargeSheetsWb());
    },
    { iterations: 10 },
  );

  bench(
    "ours STORE async — 20 sheets × 100 rows + toBufferStoreAsync",
    async () => {
      await toBufferStoreAsync(buildLargeSheetsWb());
    },
    { iterations: 10 },
  );

  bench(
    "hucre — 20 sheets × 100 rows + toBuffer",
    async () => {
      const sheets: HucreWriteSheet[] = Array.from({ length: 20 }, (_, si) => ({
        name: `Sheet${si + 1}`,
        rows: Array.from({ length: 100 }, (_, ri) => [
          `S${si + 1}R${ri + 1}`,
          ri * 10 + si,
          `Data for sheet ${si + 1} row ${ri + 1}`,
        ]),
      }));
      await hucreWriteXlsx({ sheets });
    },
    { iterations: 10 },
  );
});

// ── Large data-heavy benchmarks (100k rows) ──

const DATA_100K_ROWS = Array.from({ length: 100_000 }, (_, i) => [
  i + 1,
  `Employee ${i + 1}`,
  `Dept ${(i % 12) + 1}`,
  Math.round(30000 + Math.random() * 120000),
  `City ${(i % 50) + 1}`,
  i % 2 === 0,
  `2024-${String((i % 12) + 1).padStart(2, "0")}-${String((i % 28) + 1).padStart(2, "0")}`,
  22 + (i % 40),
  Math.round(Math.random() * 100),
  `Region ${String.fromCharCode(65 + (i % 26))}`,
  `user${i + 1}@example.com`,
  `(${String(100 + (i % 900))}) ${String(100 + (i % 900))}-${String(1000 + (i % 9000))}`,
  `${(i % 200) + 1} Main St, City ${(i % 50) + 1}`,
  `Title ${(i % 15) + 1}`,
  (i % 5) + 1,
  (Math.random() * 5).toFixed(1),
  Math.round(Math.random() * 20000),
  Math.round(Math.random() * 30000),
  Math.round(Math.random() * 80000),
  `Memo for employee ${i + 1} with additional notes`,
]);

const buildData100kWb = () =>
  new Workbook({
    worksheets: [
      {
        children: DATA_100K_ROWS.map((row) => ({
          cells: row.map((v) => ({ value: v })),
        })),
      },
    ],
  });

describe("XLSX: Large Data — 100,000 rows × 20 columns", () => {
  bench(
    "ours DEFLATE sync — 100k×20 data + toBufferSync",
    () => {
      Packer.toBufferSync(buildData100kWb());
    },
    { iterations: 3 },
  );

  bench(
    "ours STORE sync — 100k×20 data + toBufferStore",
    () => {
      toBufferStore(buildData100kWb());
    },
    { iterations: 3 },
  );

  bench(
    "ours DEFLATE async — 100k×20 data + toBuffer",
    async () => {
      await Packer.toBuffer(buildData100kWb());
    },
    { iterations: 3 },
  );

  bench(
    "ours STORE async — 100k×20 data + toBufferStoreAsync",
    async () => {
      await toBufferStoreAsync(buildData100kWb());
    },
    { iterations: 3 },
  );

  bench(
    "hucre — 100k×20 data + toBuffer",
    async () => {
      await hucreWriteXlsx({
        sheets: [{ name: "Sheet1", rows: DATA_100K_ROWS }],
      });
    },
    { iterations: 3 },
  );
});
