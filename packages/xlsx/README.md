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
import { generate } from "@office-open/xlsx";
import { writeFileSync } from "node:fs";

const buffer = await generate({
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
await generate({ worksheets: [...] });
// All STORE (no compression)
await generate({ worksheets: [...] }, { compression: { xml: 0 } });
```

**Create + toBuffer (end-to-end)**

| Scenario         | Default sync | Default async | All STORE sync | All STORE async |       hucre |
| ---------------- | -----------: | ------------: | -------------: | --------------: | ----------: |
| Simple (3 rows)  |  2,117 ops/s |   1,123 ops/s |   14,387 ops/s |    15,646 ops/s |   996 ops/s |
| Styled rows (20) |  2,174 ops/s |   1,044 ops/s |   13,924 ops/s |    13,195 ops/s | 1,012 ops/s |
| Table (10x5)     |  2,176 ops/s |   1,215 ops/s |   14,467 ops/s |    14,201 ops/s |   947 ops/s |

**Large Files — Create + toBuffer**

| Scenario                      | Default sync | Default async | All STORE sync | All STORE async |     hucre |
| ----------------------------- | -----------: | ------------: | -------------: | --------------: | --------: |
| 2000 rows + 10 images         |     69 ops/s |      66 ops/s |       75 ops/s |        77 ops/s |  42 ops/s |
| 200x10 table                  |    690 ops/s |     555 ops/s |      956 ops/s |       977 ops/s | 257 ops/s |
| 20 sheets × 100 rows + 20 img |     43 ops/s |      35 ops/s |       55 ops/s |        57 ops/s |  23 ops/s |

**Large Data — 100,000 rows × 20 columns (2M cells)**

| Scenario  | Default sync | Default async | All STORE sync | All STORE async |      hucre |
| --------- | -----------: | ------------: | -------------: | --------------: | ---------: |
| 100k × 20 |   0.75 ops/s |    0.71 ops/s |     0.86 ops/s |      0.84 ops/s | 0.37 ops/s |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
