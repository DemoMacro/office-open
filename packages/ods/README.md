# @office-open/ods

![npm version](https://img.shields.io/npm/v/@office-open/ods)
![npm downloads](https://img.shields.io/npm/dw/@office-open/ods)
![npm license](https://img.shields.io/npm/l/@office-open/ods)

> Generate and parse OpenDocument Spreadsheet (`.ods`) files from plain `WorkbookOptions` JSON — no Microsoft Office or LibreOffice required.

## Features

- 📊 **Canonical Workbook Model** — Read and write the same `WorkbookOptions` used by XLSX and XLS
- 🧮 **Worksheets** — Rows, cells, formulas, values, dimensions, and defined names
- 🎨 **Styles** — Cell, number, and dimension styles
- 🖼️ **Drawings & Charts** — Cell-anchored images and typed charts
- 🧩 **Structured Errors** — Part, path, element, and reason diagnostics

## Installation

```bash
# pnpm
pnpm add @office-open/ods

# npm
npm install @office-open/ods

# yarn
yarn add @office-open/ods

# bun
bun add @office-open/ods
```

## Quick Start

```typescript
import { generateWorkbook, parseWorkbook } from "@office-open/ods";

const ods = await generateWorkbook({
  worksheets: [{ name: "Sheet 1", rows: [{ cells: [{ value: 42 }] }] }],
});
const options = await parseWorkbook(ods);
```

## API

- `generateWorkbook(options, output?)` — generate an `.ods` package from `OdsWorkbookOptions` as `Promise` (`Uint8Array` by default); `generateWorkbookSync` covers synchronous output and `generateWorkbookStream` returns a `ReadableStream<Uint8Array>`
- `parseWorkbook(data)` — parse bytes, Base64/data-URL strings, `Blob`, or `ReadableStream` into `OdsWorkbookOptions` (the canonical `WorkbookOptions` ODF overlay) as `Promise`; `parseWorkbookSync` covers synchronous parsing (rejects `Blob` and `ReadableStream` inputs)
- `patchWorkbook({ data, commands })` — parse an ODS package, apply ordered canonical edits, and regenerate it
- `OdsParseError` — structured parse failure with part, path, and reason diagnostics

Every input is a plain JSON object (`WorkbookOptions` and its option types). The full typed API reference lives in the [documentation](https://www.office-open.com/en/getting-started/more-formats).

## Errors and Limitations

- Invalid ZIP containers, unexpected mimetypes, truncated XML, and unknown content-model elements throw `OdsParseError`.
- Generated output targets ODF 1.3.
- Patching edits the canonical model and preserves package overlays unless they are explicitly changed.

## Documentation

- [Documentation](https://www.office-open.com/en/getting-started/more-formats) — additional format guides and examples
- [Report Issues](https://github.com/DemoMacro/office-open/issues) — bug reports and feature requests

## Related Packages

- [office-open](https://www.npmjs.com/package/office-open) — all formats + CLI + AI SDK tools in one install
- [@office-open/xlsx](https://www.npmjs.com/package/@office-open/xlsx) — Excel (.xlsx) generation and parsing
- [@office-open/xls](https://www.npmjs.com/package/@office-open/xls) — legacy Excel (.xls) parsing

## License

- [MIT](https://github.com/DemoMacro/office-open/blob/main/LICENSE) © [Demo Macro](https://www.demomacro.com/)
