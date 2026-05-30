# @office-open/xlsx

![npm version](https://img.shields.io/npm/v/@office-open/xlsx)
![npm downloads](https://img.shields.io/npm/dw/@office-open/xlsx)
![npm license](https://img.shields.io/npm/l/@office-open/xlsx)

> Generate and parse .xlsx files with JS/TS with a declarative API.

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
- 📖 **Parsing** — Parse existing .xlsx files with `parseWorkbook` for round-trip workflows
- 🔧 **Template Patching** — Patch existing XLSX templates via placeholder replacement

## Installation

```bash
# Install with npm
$ npm install @office-open/xlsx

# Install with pnpm
$ pnpm add @office-open/xlsx
```

## Quick Start

```typescript
import { Workbook, Packer } from "@office-open/xlsx";
import { writeFileSync } from "node:fs";

const wb = new Workbook({
  worksheets: [
    {
      name: "Sheet1",
      children: [
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

DEFLATE = compressed (default), STORE = no compression.

**Create + toBuffer (end-to-end)**

| Scenario         | DEFLATE sync |   STORE sync | DEFLATE async |  STORE async |      hucre |
| ---------------- | -----------: | -----------: | ------------: | -----------: | ---------: |
| Simple (3 rows)  |    547 ops/s | 13,823 ops/s |     558 ops/s | 14,536 ops/s |  926 ops/s |
| Styled rows (20) |    645 ops/s | 11,275 ops/s |     644 ops/s | 11,576 ops/s |  960 ops/s |
| Table (10×5)     |    707 ops/s | 12,369 ops/s |     734 ops/s | 11,447 ops/s | 1045 ops/s |

**Large Files — Create + toBuffer**

| Scenario             | DEFLATE sync |  STORE sync | DEFLATE async | STORE async |       hucre |
| -------------------- | -----------: | ----------: | ------------: | ----------: | ----------: |
| 2000 rows            |   50.2 ops/s | 211.1 ops/s |    20.6 ops/s | 207.5 ops/s |  85.0 ops/s |
| 200×10 table         |  167.4 ops/s | 623.0 ops/s |   190.7 ops/s | 772.3 ops/s | 253.2 ops/s |
| 20 sheets × 100 rows |   87.2 ops/s | 196.4 ops/s |    93.7 ops/s | 248.6 ops/s |  87.7 ops/s |

**Large Data — 100,000 rows × 20 columns (2M cells)**

| Method        |       Speed |   Speedup |
| ------------- | ----------: | --------: |
| DEFLATE sync  | 0.274 ops/s |     0.73x |
| STORE sync    | 0.558 ops/s | **1.48x** |
| DEFLATE async | 0.281 ops/s |     0.75x |
| STORE async   | 0.694 ops/s | **1.85x** |
| hucre         | 0.376 ops/s |           |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
