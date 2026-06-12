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

**Default** = XML DEFLATE level 1 (SuperFast, matching MS Office) + media STORE. **All STORE** = no compression (`{ compression: { xml: 0 } }`). **hucre** (async only) uses `CompressionStream("deflate-raw")` when available, falls back to STORE per-entry when compression doesn't reduce size.

```typescript
// Default (matches MS Office)
await generateWorkbook(options);
// All STORE (no compression)
await generateWorkbook(options, { compression: { xml: 0 } });
```

**Create + toBuffer (end-to-end)**

| Scenario         | Default sync | Default async | All STORE sync | All STORE async |       hucre |
| ---------------- | -----------: | ------------: | -------------: | --------------: | ----------: |
| Simple (3 rows)  |  2,385 ops/s |   1,142 ops/s |   14,583 ops/s |    15,943 ops/s |   951 ops/s |
| Styled rows (20) |  2,355 ops/s |   1,195 ops/s |   16,765 ops/s |    12,852 ops/s | 1,083 ops/s |
| Table (10x5)     |  2,357 ops/s |   1,177 ops/s |   16,480 ops/s |    13,716 ops/s | 1,049 ops/s |

**Large Files — Create + toBuffer**

| Scenario                      | Default sync | Default async | All STORE sync | All STORE async |     hucre |
| ----------------------------- | -----------: | ------------: | -------------: | --------------: | --------: |
| 2000 rows + 10 images         |     79 ops/s |      75 ops/s |       89 ops/s |        93 ops/s |  44 ops/s |
| 200x10 table                  |    847 ops/s |     610 ops/s |    1,118 ops/s |     1,011 ops/s | 261 ops/s |
| 20 sheets × 100 rows + 20 img |     49 ops/s |      39 ops/s |       60 ops/s |        62 ops/s |  26 ops/s |

**Large Data — 100,000 rows × 20 columns (2M cells)**

| Scenario  | Default sync | Default async | All STORE sync | All STORE async |      hucre |
| --------- | -----------: | ------------: | -------------: | --------------: | ---------: |
| 100k × 20 |   0.84 ops/s |    0.83 ops/s |     0.90 ops/s |      0.94 ops/s | 0.39 ops/s |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
