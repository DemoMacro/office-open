# @office-open/xls

![npm version](https://img.shields.io/npm/v/@office-open/xls)
![npm downloads](https://img.shields.io/npm/dw/@office-open/xls)
![npm license](https://img.shields.io/npm/l/@office-open/xls)

> Parse legacy Microsoft Excel 97-2003 `.xls` workbooks into plain `WorkbookOptions` JSON — no Microsoft Office required.

## Features

- 📊 **Worksheet Parsing** — Read worksheets, rows, cells, values, and formulas
- ✍️ **Formatting** — Read fonts, fills, borders, number formats, and styles
- 🧮 **Formulas** — Preserve formula expressions and cached values
- 🔗 **References** — Read hyperlinks, defined names, and external workbook references
- 🖼️ **Images** — Extract embedded pictures
- 💬 **Comments** — Read cell comments and authors
- 🎛️ **Rules** — Read conditional formats
- 🔐 **Encrypted Files** — Parse password-protected workbooks with the provided password

This package is parse-only. It does not generate `.xls` files.

## Installation

```bash
# pnpm
pnpm add @office-open/xls

# npm
npm install @office-open/xls

# yarn
yarn add @office-open/xls

# bun
bun add @office-open/xls
```

## Quick Start

```typescript
import { parseWorkbook } from "@office-open/xls";
import { readFileSync } from "node:fs";

const options = parseWorkbook(new Uint8Array(readFileSync("input.xls")));
```

For an encrypted workbook, provide the password:

```typescript
const options = parseWorkbook(data, { password: "secret" });
```

## API

- `parseWorkbook(data, options?)` — read an `.xls` file into `WorkbookOptions`
- `LegacyParseOptions` — parse-time options, including `password`

The projected model follows the office-open `WorkbookOptions` shape where BIFF8 has a direct counterpart. Unsupported legacy records are skipped without aborting the workbook.

## Documentation

- [Documentation](https://www.office-open.com/en/getting-started/more-formats) — additional format guides and examples
- [Report Issues](https://github.com/DemoMacro/office-open/issues) — bug reports and feature requests

## Related Packages

- [office-open](https://www.npmjs.com/package/office-open) — all formats + CLI + AI SDK tools in one install
- [@office-open/xlsx](https://www.npmjs.com/package/@office-open/xlsx) — Excel (.xlsx) generation and parsing
- [@office-open/doc](https://www.npmjs.com/package/@office-open/doc) — legacy Word (.doc) parsing
- [@office-open/ppt](https://www.npmjs.com/package/@office-open/ppt) — legacy PowerPoint (.ppt) parsing

## License

- [MIT](https://github.com/DemoMacro/office-open/blob/main/LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
