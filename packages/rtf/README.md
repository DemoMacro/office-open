# @office-open/rtf

![npm version](https://img.shields.io/npm/v/@office-open/rtf)
![npm downloads](https://img.shields.io/npm/dw/@office-open/rtf)
![npm license](https://img.shields.io/npm/l/@office-open/rtf)

> Generate and parse Rich Text Format (RTF) documents with plain `DocumentOptions` JSON — no Microsoft Office required.

## Features

- ✍️ **Rich Text** — Parse text runs, character formatting, fonts, colors, Unicode, and hexadecimal escapes
- 📑 **Paragraphs** — Read alignment, spacing, indentation, tabs, borders, and shading
- 📊 **Tables** — Read flat and nested table structure and cell formatting
- 🎨 **Styles** — Read character, paragraph, and table style definitions
- 📋 **Lists** — Read numbering definitions, list levels, and overrides
- 🖼️ **Shapes** — Capture embedded shape instructions
- 🔗 **Hyperlinks** — Read field-based and direct hyperlinks
- 📝 **Fields & Metadata** — Read fields, titles, subjects, authors, keywords, dates, and categories
- 📃 **Sections** — Read page geometry, headers, and footers
- ✍️ **Rich Text Writer** — Generate RTF text, formatting, tables, fields, pictures, and shapes

## Installation

```bash
# pnpm
pnpm add @office-open/rtf

# npm
npm install @office-open/rtf

# yarn
yarn add @office-open/rtf

# bun
bun add @office-open/rtf
```

## Quick Start

```typescript
import { generateRtf, parseRtf } from "@office-open/rtf";
import { readFileSync } from "node:fs";

const options = parseRtf(readFileSync("input.rtf", "utf8"));
const rtf = new TextDecoder().decode(generateRtf(options));
```

## API

- `parseRtf(source)` — read RTF source into `RtfDocumentOptions`
- `generateRtf(options)` — write canonical `DocumentOptions` to RTF bytes
- `RtfDocumentOptions` — the projected Word document model
- `RtfParseError` — thrown for malformed RTF

The projected model follows the office-open `DocumentOptions` shape. Unrecognized destinations are skipped without aborting the document. Generation rejects fields that cannot be represented in RTF with `RtfGenerateError` instead of dropping them silently.

## Documentation

- [Documentation](https://www.office-open.com/en/getting-started/more-formats) — additional format guides and examples
- [Report Issues](https://github.com/DemoMacro/office-open/issues) — bug reports and feature requests

## Related Packages

- [office-open](https://www.npmjs.com/package/office-open) — all formats + CLI + AI SDK tools in one install
- [@office-open/docx](https://www.npmjs.com/package/@office-open/docx) — Word (.docx) generation and parsing
- [@office-open/doc](https://www.npmjs.com/package/@office-open/doc) — legacy Word (.doc) parsing
- [@office-open/odt](https://www.npmjs.com/package/@office-open/odt) — OpenDocument Text generation and parsing

## License

- [MIT](https://github.com/DemoMacro/office-open/blob/main/LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
