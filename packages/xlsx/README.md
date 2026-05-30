# @office-open/xlsx

![npm version](https://img.shields.io/npm/v/@office-open/xlsx)
![npm downloads](https://img.shields.io/npm/dw/@office-open/xlsx)
![npm license](https://img.shields.io/npm/l/@office-open/xlsx)

> Generate and parse .xlsx files with JS/TS with a declarative API.

## Features

- **Workbook Generation** — Create spreadsheets with multiple worksheets
- **Cell Data** — Strings, numbers, booleans, dates, and inline strings
- **Styles** — Fonts, fills, borders, alignment, and number formats via index-based style system
- **Merged Cells** — Merge cell ranges across rows and columns
- **Column Width & Row Height** — Custom column widths and row heights with hiding support
- **Freeze Panes** — Freeze rows and/or columns for scrollable headers
- **Auto Filter** — Add auto-filter dropdowns to column headers
- **Images** — Embed PNG and JPEG images anchored to cells
- **Charts** — Bar, line, pie, area, and scatter charts with customization
- **Data Validation** — List, whole number, decimal, date, and custom validations
- **Conditional Formatting** — Cell value-based rules with formatting
- **Parsing** — Parse existing .xlsx files with `parseWorkbook` for round-trip workflows
- **Template Patching** — Patch existing XLSX templates via placeholder replacement

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

Performance comparison against [hucre](https://github.com/nicolo-ribaudo/hucre) (0.6.0) (higher hz is better, Windows 11 / Node 24).

Both libraries use DEFLATE compression.

**Create + toBuffer (end-to-end)**

Both libraries use DEFLATE compression.

| Scenario                    | @office-open/xlsx (sync) | @office-open/xlsx (async) |       hucre | Speedup (sync) | Speedup (async) |
| --------------------------- | -----------------------: | ------------------------: | ----------: | -------------: | --------------: |
| Simple (3 rows) + toBuffer  |                697 ops/s |                 908 ops/s | 1,006 ops/s |          0.69x |           0.90x |
| Styled rows (20) + toBuffer |                686 ops/s |               1,144 ops/s |   848 ops/s |          0.81x |           1.35x |
| Table (10x5) + toBuffer     |                793 ops/s |               1,142 ops/s |   877 ops/s |          0.90x |           1.30x |

**Large Files — Create + toBuffer**

| Scenario             | @office-open/xlsx (sync) | @office-open/xlsx (async) |      hucre | Speedup (sync) | Speedup (async) |
| -------------------- | -----------------------: | ------------------------: | ---------: | -------------: | --------------: |
| 2000 rows            |               46.7 ops/s |                14.0 ops/s | 77.0 ops/s |          0.61x |           0.18x |
| 200×10 table         |              141.4 ops/s |               180.9 ops/s |  230 ops/s |          0.61x |           0.79x |
| 20 sheets × 100 rows |               70.3 ops/s |                74.7 ops/s | 72.2 ops/s |          0.97x |           1.04x |

**Large File (~100MB) — Mixed Content**

2000 rows + 200 unique images (500KB each). Speedup is vs hucre.

| Method                    |      Speed |   Speedup |
| ------------------------- | ---------: | --------: |
| @office-open/xlsx (sync)  | 1.92 ops/s |     1.17x |
| @office-open/xlsx (async) | 3.58 ops/s | **2.18x** |
| hucre                     | 1.65 ops/s |           |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
