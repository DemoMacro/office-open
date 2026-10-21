# @office-open/odt

![npm version](https://img.shields.io/npm/v/@office-open/odt)
![npm downloads](https://img.shields.io/npm/dw/@office-open/odt)
![npm license](https://img.shields.io/npm/l/@office-open/odt)

> Generate and parse OpenDocument Text (`.odt`) files from plain `DocumentOptions` JSON — no Microsoft Office or LibreOffice required.

## Features

- 📄 **Canonical Text Model** — Read and write the same `DocumentOptions` used by DOCX, DOC, and RTF
- ✍️ **Rich Content** — Paragraphs, runs, lists, tables, bookmarks, fields, notes, and drawings
- 🎨 **Styles** — Automatic, common, and page styles
- 🖼️ **Media & Charts** — Embedded images and typed chart projection
- 🧩 **Structured Errors** — Part, path, element, and reason diagnostics

## Installation

```bash
# pnpm
pnpm add @office-open/odt

# npm
npm install @office-open/odt

# yarn
yarn add @office-open/odt

# bun
bun add @office-open/odt
```

## Quick Start

```typescript
import { generateDocument, parseDocument } from "@office-open/odt";

const odt = await generateDocument({
  sections: [{ children: [{ paragraph: "Hello ODT" }] }],
});
const options = await parseDocument(odt);
```

## API

- `generateDocument(options, output?)` — generate an `.odt` package from `OdtDocumentOptions` as `Promise` (`Uint8Array` by default); `generateDocumentSync` covers synchronous output and `generateDocumentStream` returns a `ReadableStream<Uint8Array>`
- `parseDocument(data)` — parse bytes, Base64/data-URL strings, `Blob`, or `ReadableStream` into `OdtDocumentOptions` (the canonical `DocumentOptions` ODF overlay) as `Promise`; `parseDocumentSync` covers synchronous parsing (rejects `Blob` and `ReadableStream` inputs)
- `patchDocument({ data, commands })` — parse an ODT package, apply ordered canonical edits, and regenerate it
- `OdtParseError` — structured parse failure with part, path, and reason diagnostics

Every input is a plain JSON object (`DocumentOptions` and its option types). The full typed API reference lives in the [documentation](https://www.office-open.com/en/getting-started/more-formats).

## Errors and Limitations

- Invalid ZIP containers, unexpected mimetypes, truncated XML, and unknown content-model elements throw `OdtParseError`.
- Generated output targets ODF 1.3.
- Patching edits the canonical model and preserves package overlays unless they are explicitly changed.

## Documentation

- [Documentation](https://www.office-open.com/en/getting-started/more-formats) — additional format guides and examples
- [Report Issues](https://github.com/DemoMacro/office-open/issues) — bug reports and feature requests

## Related Packages

- [office-open](https://www.npmjs.com/package/office-open) — all formats + CLI + AI SDK tools in one install
- [@office-open/docx](https://www.npmjs.com/package/@office-open/docx) — Word (.docx) generation and parsing
- [@office-open/doc](https://www.npmjs.com/package/@office-open/doc) — legacy Word (.doc) parsing
- [@office-open/rtf](https://www.npmjs.com/package/@office-open/rtf) — Rich Text Format generation and parsing

## License

- [MIT](https://github.com/DemoMacro/office-open/blob/main/LICENSE) © [Demo Macro](https://www.demomacro.com/)
