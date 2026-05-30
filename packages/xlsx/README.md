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

| Scenario         | DEFLATE sync |   STORE sync | DEFLATE async |  STORE async |     hucre |
| ---------------- | -----------: | -----------: | ------------: | -----------: | --------: |
| Simple (3 rows)  |    457 ops/s | 10,016 ops/s |     500 ops/s | 11,637 ops/s | 689 ops/s |
| Styled rows (20) |    472 ops/s | 10,504 ops/s |     627 ops/s |  9,479 ops/s | 862 ops/s |
| Table (10×5)     |    675 ops/s | 10,518 ops/s |     681 ops/s | 10,241 ops/s | 877 ops/s |

**Large Files — Create + toBuffer**

| Scenario             | DEFLATE sync |  STORE sync | DEFLATE async | STORE async |       hucre |
| -------------------- | -----------: | ----------: | ------------: | ----------: | ----------: |
| 2000 rows            |   47.3 ops/s | 164.9 ops/s |    19.0 ops/s | 162.9 ops/s |  82.8 ops/s |
| 200×10 table         |  148.2 ops/s | 510.5 ops/s |   160.7 ops/s | 691.7 ops/s | 226.4 ops/s |
| 20 sheets × 100 rows |   81.3 ops/s | 179.5 ops/s |    87.8 ops/s | 223.5 ops/s |  79.9 ops/s |

**Large Data — 100,000 rows × 20 columns (2M cells)**

| Method        |       Speed |   Speedup |
| ------------- | ----------: | --------: |
| DEFLATE sync  | 0.250 ops/s |     0.68x |
| STORE sync    | 0.516 ops/s | **1.40x** |
| DEFLATE async | 0.270 ops/s |     0.73x |
| STORE async   | 0.664 ops/s | **1.80x** |
| hucre         | 0.369 ops/s |           |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
