# @office-open/xlsx

![npm version](https://img.shields.io/npm/v/@office-open/xlsx)
![npm downloads](https://img.shields.io/npm/dw/@office-open/xlsx)
![npm license](https://img.shields.io/npm/l/@office-open/xlsx)

> Generate and parse .xlsx spreadsheets with a declarative TypeScript API. Works in Node.js and browsers.

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
import { Workbook, Packer } from "@office-open/xlsx";
import { writeFileSync } from "node:fs";

const wb = new Workbook({
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

const buffer = await Packer.toBuffer(wb);
writeFileSync("workbook.xlsx", buffer);
```

## Examples

Check the [demo folder](./demo) for working examples covering every feature.

## Benchmark

Performance comparison against [hucre](https://github.com/nicolo-ribaudo/hucre) (0.6.0) (higher ops/s is better, Windows 11 / Node 24).

**Default** = XML DEFLATE level 1 (SuperFast, matching MS Office) + media STORE. **All STORE** = no compression (`{ compression: { xml: 0 } }`). **hucre** (async only) uses `CompressionStream("deflate-raw")` when available, falls back to STORE per-entry when compression doesn't reduce size.

```typescript
// Default (matches MS Office)
await Packer.toBuffer(wb);
// All STORE (no compression)
await Packer.toBuffer(wb, { compression: { xml: 0 } });
```

**Create + toBuffer (end-to-end)**

| Scenario         | Default sync | Default async | All STORE sync | All STORE async |       hucre |
| ---------------- | -----------: | ------------: | -------------: | --------------: | ----------: |
| Simple (3 rows)  |  3,528 ops/s |   1,266 ops/s |   17,759 ops/s |    17,193 ops/s |   949 ops/s |
| Styled rows (20) |  3,403 ops/s |   1,289 ops/s |   16,516 ops/s |    13,811 ops/s |   971 ops/s |
| Table (10x5)     |  3,449 ops/s |   1,414 ops/s |   16,036 ops/s |    13,475 ops/s | 1,031 ops/s |

**Large Files — Create + toBuffer**

| Scenario                      | Default sync | Default async | All STORE sync | All STORE async |      hucre |
| ----------------------------- | -----------: | ------------: | -------------: | --------------: | ---------: |
| 2000 rows + 10 images         |   87.5 ops/s |    79.7 ops/s |    101.0 ops/s |     101.2 ops/s | 46.4 ops/s |
| 200x10 table                  |    754 ops/s |     517 ops/s |    1,049 ops/s |       963 ops/s |  226 ops/s |
| 20 sheets x 100 rows + 20 img |   60.9 ops/s |    49.9 ops/s |     76.8 ops/s |      77.5 ops/s | 25.3 ops/s |

**Large Data — 100,000 rows x 20 columns (2M cells)**

| Method    |      Speed |  Speedup |
| --------- | ---------: | -------: |
| All STORE | 0.74 ops/s | **1.9x** |
| Default   | 0.66 ops/s |     1.7x |
| hucre     | 0.40 ops/s |          |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
