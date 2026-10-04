# @office-open/doc

![npm version](https://img.shields.io/npm/v/@office-open/doc)
![npm downloads](https://img.shields.io/npm/dw/@office-open/doc)
![npm license](https://img.shields.io/npm/l/@office-open/doc)

> Parse legacy Microsoft Word 97-2003 `.doc` documents into plain `DocumentOptions` JSON — no Microsoft Office required.

## Features

- 📄 **Document Parsing** — Read body content, sections, headers, and footers
- ✍️ **Rich Text** — Preserve text runs and character formatting
- 📊 **Tables** — Read table structure, rows, cells, and merged cells
- 🖼️ **Images** — Extract embedded pictures
- 🔗 **Hyperlinks** — Resolve internal and external link targets
- 🔖 **Bookmarks** — Read bookmark starts and ends
- 📝 **Fields** — Capture field instructions and cached results
- 🕘 **Revisions** — Read insertion and deletion revision ranges
- 🔐 **Encrypted Files** — Parse password-protected documents with the provided password

This package is parse-only. It does not generate `.doc` files.

## Installation

```bash
# pnpm
pnpm add @office-open/doc

# npm
npm install @office-open/doc

# yarn
yarn add @office-open/doc

# bun
bun add @office-open/doc
```

## Quick Start

```typescript
import { parseDocument } from "@office-open/doc";
import { readFileSync } from "node:fs";

const options = parseDocument(new Uint8Array(readFileSync("input.doc")));
```

For an encrypted document, provide the password:

```typescript
const options = parseDocument(data, { password: "secret" });
```

## API

- `parseDocument(data, options?)` — read a `.doc` file into `DocumentOptions`
- `DocParseOptions` — parse-time options, including `password`
- `DocParseError` — thrown for invalid, truncated, unsupported, or encrypted input
- `DocParseError` — thrown for invalid, truncated, unsupported, or encrypted input

The projected model follows the office-open `DocumentOptions` shape where the legacy format has a direct counterpart. Unsupported legacy records are skipped without aborting the document.

## Documentation

- [Documentation](https://www.office-open.com/en/getting-started/more-formats) — additional format guides and examples
- [Report Issues](https://github.com/DemoMacro/office-open/issues) — bug reports and feature requests

## Related Packages

- [office-open](https://www.npmjs.com/package/office-open) — all formats + CLI + AI SDK tools in one install
- [@office-open/docx](https://www.npmjs.com/package/@office-open/docx) — Word (.docx) generation and parsing
- [@office-open/xls](https://www.npmjs.com/package/@office-open/xls) — legacy Excel (.xls) parsing
- [@office-open/ppt](https://www.npmjs.com/package/@office-open/ppt) — legacy PowerPoint (.ppt) parsing

## License

- [MIT](https://github.com/DemoMacro/office-open/blob/main/LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
