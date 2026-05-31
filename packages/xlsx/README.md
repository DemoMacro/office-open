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
| Simple (3 rows)  |    551 ops/s | 12,909 ops/s |     597 ops/s | 14,019 ops/s | 773 ops/s |
| Styled rows (20) |    759 ops/s | 11,510 ops/s |     857 ops/s | 11,225 ops/s | 877 ops/s |
| Table (10×5)     |    868 ops/s | 12,101 ops/s |     937 ops/s | 10,705 ops/s | 949 ops/s |

**Large Files — Create + toBuffer**

| Scenario             | DEFLATE sync |  STORE sync | DEFLATE async | STORE async |       hucre |
| -------------------- | -----------: | ----------: | ------------: | ----------: | ----------: |
| 2000 rows            |   58.0 ops/s | 192.5 ops/s |    21.2 ops/s | 201.5 ops/s |  80.3 ops/s |
| 200×10 table         |  178.4 ops/s | 589.1 ops/s |   196.6 ops/s | 811.0 ops/s | 213.7 ops/s |
| 20 sheets × 100 rows |   89.8 ops/s | 192.9 ops/s |    93.4 ops/s | 243.5 ops/s |  83.1 ops/s |

**Large Data — 100,000 rows × 20 columns (2M cells)**

| Method        |       Speed |   Speedup |
| ------------- | ----------: | --------: |
| DEFLATE sync  | 0.290 ops/s |     0.75x |
| STORE sync    | 0.537 ops/s | **1.39x** |
| DEFLATE async | 0.322 ops/s |     0.83x |
| STORE async   | 0.649 ops/s | **1.68x** |
| hucre         | 0.386 ops/s |           |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
