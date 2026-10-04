# @office-open/ppt

![npm version](https://img.shields.io/npm/v/@office-open/ppt)
![npm downloads](https://img.shields.io/npm/dw/@office-open/ppt)
![npm license](https://img.shields.io/npm/l/@office-open/ppt)

> Parse legacy Microsoft PowerPoint 97-2003 `.ppt` presentations into plain `PresentationOptions` JSON — no Microsoft Office required.

## Features

- 📄 **Presentation Parsing** — Read slide size, slides, masters, and speaker notes
- ✍️ **Text & Paragraphs** — Preserve text runs, formatting, lists, and alignment
- 🧩 **Shapes** — Read shapes, text boxes, groups, and pictures
- 📊 **Tables** — Read table structure, rows, cells, and cell formatting
- 🔗 **Hyperlinks** — Resolve internal and external link targets
- 🎬 **Animations** — Read legacy animation entries
- 🎨 **Formatting** — Read fills, lines, and transforms
- 🔐 **Encrypted Files** — Parse password-protected presentations with the provided password

This package is parse-only. It does not generate `.ppt` files.

## Installation

```bash
# pnpm
pnpm add @office-open/ppt

# npm
npm install @office-open/ppt

# yarn
yarn add @office-open/ppt

# bun
bun add @office-open/ppt
```

## Quick Start

```typescript
import { parsePresentation } from "@office-open/ppt";
import { readFileSync } from "node:fs";

const options = parsePresentation(new Uint8Array(readFileSync("input.ppt")));
```

For an encrypted presentation, provide the password:

```typescript
const options = parsePresentation(data, { password: "secret" });
```

## API

- `parsePresentation(data, options?)` — read a `.ppt` file into `PresentationOptions`
- `PptParseOptions` — parse-time options, including `password`
- `PptParseError` — thrown for invalid, truncated, unsupported, or encrypted input, with canonical `context` fields

The projected model follows the office-open `PresentationOptions` shape where the legacy binary format has a direct counterpart. Unsupported legacy records are skipped without aborting the presentation.

## Documentation

- [Documentation](https://www.office-open.com/en/getting-started/more-formats) — additional format guides and examples
- [Report Issues](https://github.com/DemoMacro/office-open/issues) — bug reports and feature requests

## Related Packages

- [office-open](https://www.npmjs.com/package/office-open) — all formats + CLI + AI SDK tools in one install
- [@office-open/pptx](https://www.npmjs.com/package/@office-open/pptx) — PowerPoint (.pptx) generation and parsing
- [@office-open/doc](https://www.npmjs.com/package/@office-open/doc) — legacy Word (.doc) parsing
- [@office-open/xls](https://www.npmjs.com/package/@office-open/xls) — legacy Excel (.xls) parsing

## License

- [MIT](https://github.com/DemoMacro/office-open/blob/main/LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
