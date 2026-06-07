import { writeXlsx as hucreWriteXlsx } from "hucre";
import type { WriteSheet as HucreWriteSheet } from "hucre";
import { bench, describe } from "vite-plus/test";

import { generate, generateSync } from "./index";

// Bench modes:
//   "ours default"  = XML DEFLATE level 1 (SuperFast, MS Office), media STORE — no options passed.
//   "ours all-store" = all entries STORE — { compression: { xml: 0 } }.
//
// hucre: async only (writeXlsx). Per-entry compression:
// XML → DEFLATE (auto fallback to STORE), images → STORE (explicit compress: false).

// ── Image generation ──

const makeImage = (seed: number, sizeKB: number): Uint8Array => {
  const size = sizeKB * 1024;
  const buf = new Uint8Array(size);
  for (let i = 0; i < size; i++) buf[i] = (i * 7 + seed * 13 + 37) & 0xff;
  return buf;
};

const LARGE_IMAGES = Array.from({ length: 20 }, (_, i) => makeImage(i, 500));

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

const SIMPLE_OPTS = {
  worksheets: [
    {
      rows: SIMPLE_ROWS.map((row) => ({
        cells: row.map((v) => ({ value: v })),
      })),
    },
  ],
};

const STYLED_OPTS = {
  worksheets: [
    {
      rows: STYLED_ROWS.map((r) => ({
        cells: [{ value: r.name }, { value: r.score }, { value: r.active }],
      })),
    },
  ],
};

const TABLE_OPTS = {
  worksheets: [
    {
      rows: TABLE_ROWS.map((row) => ({
        cells: row.map((v) => ({ value: v })),
      })),
    },
  ],
};

// ── Benchmarks ──

describe("XLSX: Create + toBuffer", () => {
  bench(
    "ours default sync — simple + toBufferSync",
    () => {
      generateSync(SIMPLE_OPTS);
    },
    { iterations: 50 },
  );

  bench(
    "ours all-store sync — simple + toBufferStore",
    () => {
      generateSync(SIMPLE_OPTS, { compression: { xml: 0 } });
    },
    { iterations: 50 },
  );

  bench(
    "ours default async — simple + toBuffer",
    async () => {
      await generate(SIMPLE_OPTS);
    },
    { iterations: 50 },
  );

  bench(
    "ours all-store async — simple + toBufferStoreAsync",
    async () => {
      await generate(SIMPLE_OPTS, { compression: { xml: 0 } });
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
    "ours default sync — styled rows (20) + toBufferSync",
    () => {
      generateSync(STYLED_OPTS);
    },
    { iterations: 50 },
  );

  bench(
    "ours all-store sync — styled rows (20) + toBufferStore",
    () => {
      generateSync(STYLED_OPTS, { compression: { xml: 0 } });
    },
    { iterations: 50 },
  );

  bench(
    "ours default async — styled rows (20) + toBuffer",
    async () => {
      await generate(STYLED_OPTS);
    },
    { iterations: 50 },
  );

  bench(
    "ours all-store async — styled rows (20) + toBufferStoreAsync",
    async () => {
      await generate(STYLED_OPTS, { compression: { xml: 0 } });
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
    "ours default sync — table (10x5) + toBufferSync",
    () => {
      generateSync(TABLE_OPTS);
    },
    { iterations: 50 },
  );

  bench(
    "ours all-store sync — table (10x5) + toBufferStore",
    () => {
      generateSync(TABLE_OPTS, { compression: { xml: 0 } });
    },
    { iterations: 50 },
  );

  bench(
    "ours default async — table (10x5) + toBuffer",
    async () => {
      await generate(TABLE_OPTS);
    },
    { iterations: 50 },
  );

  bench(
    "ours all-store async — table (10x5) + toBufferStoreAsync",
    async () => {
      await generate(TABLE_OPTS, { compression: { xml: 0 } });
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

const LARGE_ROWS_OPTS = {
  worksheets: [
    {
      rows: LARGE_ROWS.map((row) => ({
        cells: row.map((v) => ({ value: v })),
      })),
      images: Array.from({ length: 10 }, (_, i) => ({
        data: LARGE_IMAGES[i],
        type: "jpeg" as const,
        col: 5,
        row: i * 200,
      })),
    },
  ],
};

const LARGE_TABLE_OPTS = {
  worksheets: [
    {
      rows: LARGE_TABLE_ROWS.map((row) => ({
        cells: row.map((v) => ({ value: v })),
      })),
    },
  ],
};

const LARGE_SHEETS_OPTS = {
  worksheets: Array.from({ length: 20 }, (_, si) => ({
    name: `Sheet${si + 1}`,
    rows: Array.from({ length: 100 }, (_, ri) => ({
      cells: [
        { value: `S${si + 1}R${ri + 1}` },
        { value: ri * 10 + si },
        { value: `Data for sheet ${si + 1} row ${ri + 1}` },
      ],
    })),
    images: [
      {
        data: LARGE_IMAGES[si % LARGE_IMAGES.length],
        type: "jpeg" as const,
        col: 3,
        row: 50,
      },
    ],
  })),
};

describe("XLSX: Large Files — Create + toBuffer", () => {
  bench(
    "ours default sync — 2000 rows + 10 img + toBufferSync",
    () => {
      generateSync(LARGE_ROWS_OPTS);
    },
    { iterations: 10 },
  );

  bench(
    "ours all-store sync — 2000 rows + 10 img + toBufferStore",
    () => {
      generateSync(LARGE_ROWS_OPTS, { compression: { xml: 0 } });
    },
    { iterations: 10 },
  );

  bench(
    "ours default async — 2000 rows + 10 img + toBuffer",
    async () => {
      await generate(LARGE_ROWS_OPTS);
    },
    { iterations: 10 },
  );

  bench(
    "ours all-store async — 2000 rows + 10 img + toBufferStoreAsync",
    async () => {
      await generate(LARGE_ROWS_OPTS, { compression: { xml: 0 } });
    },
    { iterations: 10 },
  );

  bench(
    "hucre — 2000 rows + 10 img + toBuffer",
    async () => {
      await hucreWriteXlsx({
        sheets: [
          {
            name: "Sheet1",
            rows: LARGE_ROWS,
            images: Array.from({ length: 10 }, (_, i) => ({
              data: LARGE_IMAGES[i],
              type: "jpeg" as const,
              anchor: { from: { row: i * 200, col: 5 } },
              altText: `Image ${i + 1}`,
            })),
          },
        ],
      });
    },
    { iterations: 10 },
  );

  bench(
    "ours default sync — 200x10 table + toBufferSync",
    () => {
      generateSync(LARGE_TABLE_OPTS);
    },
    { iterations: 10 },
  );

  bench(
    "ours all-store sync — 200x10 table + toBufferStore",
    () => {
      generateSync(LARGE_TABLE_OPTS, { compression: { xml: 0 } });
    },
    { iterations: 10 },
  );

  bench(
    "ours default async — 200x10 table + toBuffer",
    async () => {
      await generate(LARGE_TABLE_OPTS);
    },
    { iterations: 10 },
  );

  bench(
    "ours all-store async — 200x10 table + toBufferStoreAsync",
    async () => {
      await generate(LARGE_TABLE_OPTS, { compression: { xml: 0 } });
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
    "ours default sync — 20 sheets × 100 rows + 20 img + toBufferSync",
    () => {
      generateSync(LARGE_SHEETS_OPTS);
    },
    { iterations: 10 },
  );

  bench(
    "ours all-store sync — 20 sheets × 100 rows + 20 img + toBufferStore",
    () => {
      generateSync(LARGE_SHEETS_OPTS, { compression: { xml: 0 } });
    },
    { iterations: 10 },
  );

  bench(
    "ours default async — 20 sheets × 100 rows + 20 img + toBuffer",
    async () => {
      await generate(LARGE_SHEETS_OPTS);
    },
    { iterations: 10 },
  );

  bench(
    "ours all-store async — 20 sheets × 100 rows + 20 img + toBufferStoreAsync",
    async () => {
      await generate(LARGE_SHEETS_OPTS, { compression: { xml: 0 } });
    },
    { iterations: 10 },
  );

  bench(
    "hucre — 20 sheets × 100 rows + 20 img + toBuffer",
    async () => {
      const sheets: HucreWriteSheet[] = Array.from({ length: 20 }, (_, si) => ({
        name: `Sheet${si + 1}`,
        rows: Array.from({ length: 100 }, (_, ri) => [
          `S${si + 1}R${ri + 1}`,
          ri * 10 + si,
          `Data for sheet ${si + 1} row ${ri + 1}`,
        ]),
        images: [
          {
            data: LARGE_IMAGES[si % LARGE_IMAGES.length],
            type: "jpeg" as const,
            anchor: { from: { row: 50, col: 3 } },
            altText: `Image for sheet ${si + 1}`,
          },
        ],
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

const DATA_100K_OPTS = {
  worksheets: [
    {
      rows: DATA_100K_ROWS.map((row) => ({
        cells: row.map((v) => ({ value: v })),
      })),
    },
  ],
};

describe("XLSX: Large Data — 100,000 rows × 20 columns", () => {
  bench(
    "ours default sync — 100k×20 data + toBufferSync",
    () => {
      generateSync(DATA_100K_OPTS);
    },
    { iterations: 3 },
  );

  bench(
    "ours all-store sync — 100k×20 data + toBufferStore",
    () => {
      generateSync(DATA_100K_OPTS, { compression: { xml: 0 } });
    },
    { iterations: 3 },
  );

  bench(
    "ours default async — 100k×20 data + toBuffer",
    async () => {
      await generate(DATA_100K_OPTS);
    },
    { iterations: 3 },
  );

  bench(
    "ours all-store async — 100k×20 data + toBufferStoreAsync",
    async () => {
      await generate(DATA_100K_OPTS, { compression: { xml: 0 } });
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
