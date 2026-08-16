# @office-open/xlsx

![npm version](https://img.shields.io/npm/v/@office-open/xlsx)
![npm downloads](https://img.shields.io/npm/dw/@office-open/xlsx)
![npm license](https://img.shields.io/npm/l/@office-open/xlsx)

> Create Excel spreadsheets (.xlsx) in TypeScript and JavaScript — generate, parse, and patch from plain JSON, no Microsoft Office required. Built for AI agents and hand-written code alike; runs in Node.js, browsers, Deno, and Bun.

## Features

- 📗 **Workbook Generation** — Create spreadsheets with multiple worksheets
- 📊 **Cell Data** — Strings, numbers, booleans, dates, and inline strings
- 🎨 **Styles** — Fonts, fills, borders, alignment, and number formats via index-based style system
- 🔀 **Merged Cells** — Merge cell ranges across rows and columns
- 📏 **Column Width & Row Height** — Custom column widths and row heights with hiding support
- ❄️ **Freeze Panes** — Freeze rows and/or columns for scrollable headers
- 🔽 **Auto Filter** — Add auto-filter dropdowns to column headers
- 🖼️ **Images** — Embed PNG and JPEG images anchored to cells
- 📈 **Charts** — Bar, line, pie, area, and scatter charts with customization
- ✅ **Data Validation** — List, whole number, decimal, date, and custom validations
- 🎯 **Conditional Formatting** — Cell value-based rules with formatting
- 📊 **Pivot Tables** — Create pivot tables with various aggregation functions
- 💬 **Comments** — Cell comments with author tracking and rich text support
- 🔒 **Sheet & Workbook Protection** — Password-protect worksheets and workbook structure
- 📖 **Parsing** — Parse existing .xlsx files with `parseWorkbook` for round-trip workflows
- 🔧 **Template Patching** — Patch existing XLSX templates via placeholder replacement

## Installation

```bash
# pnpm
pnpm add @office-open/xlsx

# npm
npm install @office-open/xlsx

# yarn
yarn add @office-open/xlsx

# bun
bun add @office-open/xlsx
```

## Quick Start

```typescript
import { generateWorkbook } from "@office-open/xlsx";
import { writeFileSync } from "node:fs";

const buffer = await generateWorkbook({
  worksheets: [
    {
      name: "Sheet1",
      rows: [
        { cells: [{ value: "Name" }, { value: "Score" }] },
        { cells: [{ value: "Alice" }, { value: 95 }] },
        { cells: [{ value: "Bob" }, { value: 87 }] },
      ],
    },
  ],
});

writeFileSync("workbook.xlsx", buffer);
```

## Examples

Check the [demo folder](./demo) for working examples covering every feature.

## Benchmark

Performance comparison against [hucre](https://github.com/nicolo-ribaudo/hucre) (higher ops/s is better, Windows 11 / Node 24).

**Default** = XML DEFLATE level 1 (SuperFast); media is split by type, matching MS Office Excel — already-compressed formats (PNG/JPEG/GIF) are STOREd, the rest (EMF/WMF/BMP/TIFF/…) use DEFLATE level 6 / Normal (verified on a real MS Office file). **All STORE** = no compression (`{ compression: { xml: 0, media: 0 } }`). **hucre** (async only) uses `CompressionStream("deflate-raw")` when available, falls back to STORE per-entry when compression doesn't reduce size.

```typescript
// Default (matches MS Office)
await generateWorkbook(options);
// All STORE (no compression)
await generateWorkbook(options, { compression: { xml: 0, media: 0 } });
// Stream as ReadableStream<Uint8Array> (pipe to a file / HTTP response)
generateWorkbookStream(options);
```

**Create + toBuffer / toStream (end-to-end)**

| Scenario         | Default sync | Default async | All STORE sync | All STORE async | Default stream |     hucre |
| ---------------- | -----------: | ------------: | -------------: | --------------: | -------------: | --------: |
| Simple (3 rows)  |  1,923 ops/s |   1,115 ops/s |   18,126 ops/s |    16,147 ops/s |     22.0 ops/s | 782 ops/s |
| Styled rows (20) |  2,120 ops/s |   1,199 ops/s |   16,476 ops/s |    16,039 ops/s |     21.5 ops/s | 797 ops/s |
| Table (10x5)     |  1,847 ops/s |   1,105 ops/s |   16,181 ops/s |    16,204 ops/s |     21.9 ops/s | 875 ops/s |

**Large Files — Create + toBuffer / toStream**

| Scenario                      | Default sync | Default async | All STORE sync | All STORE async | Default stream |      hucre |
| ----------------------------- | -----------: | ------------: | -------------: | --------------: | -------------: | ---------: |
| 2000 rows + 10 images         |    131 ops/s |     138 ops/s |      135 ops/s |       140 ops/s |     12.1 ops/s | 42.2 ops/s |
| 200x10 table                  |    834 ops/s |     610 ops/s |    1,357 ops/s |     1,337 ops/s |     19.9 ops/s |  242 ops/s |
| 20 sheets × 100 rows + 20 img |   71.1 ops/s |    53.3 ops/s |     97.3 ops/s |       101 ops/s |     2.26 ops/s | 24.3 ops/s |

**Large Data — 100,000 rows × 20 columns (2M cells)**

| Scenario  | Default sync | Default async | All STORE sync | All STORE async | Default stream |      hucre |
| --------- | -----------: | ------------: | -------------: | --------------: | -------------: | ---------: |
| 100k × 20 |   0.92 ops/s |    0.90 ops/s |     1.16 ops/s |      1.03 ops/s |     0.96 ops/s | 0.38 ops/s |

**Stream column** = `generateWorkbookStream` (default compression, fully drained). Streaming trades throughput for pipeability: each part is compressed in a Web Worker as it is produced, so a fixed per-part worker handoff dominates small workbooks. Plain-data workbooks stream through a constant-memory path — worksheet rows serialize chunk-wise with inline strings, so the full worksheet XML and the archive never coexist (at 1M rows × 3 cols: peak RSS +177 MB streamed vs +810 MB buffered, and ~17% faster; at 100k × 20 it is also the fastest mode). Scenarios with images fall back to the full-memory stream.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
