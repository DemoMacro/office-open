# @office-open/odt

![npm version](https://img.shields.io/npm/v/@office-open/odt)
![npm downloads](https://img.shields.io/npm/dw/@office-open/odt)
![npm license](https://img.shields.io/npm/l/@office-open/odt)

> Generate and parse OpenDocument Text `.odt` files using canonical `DocumentOptions` JSON.

## Features

- 📄 **Canonical Text Model** — Read and write the same `DocumentOptions` used by DOCX, DOC, and RTF
- ✍️ **Rich Content** — Paragraphs, runs, lists, tables, bookmarks, fields, notes, and drawings
- 🎨 **Styles** — Automatic, common, and page styles
- 🖼️ **Media & Charts** — Embedded images and typed chart projection
- 🧩 **Structured Errors** — Part, path, element, and reason diagnostics

## Installation

```bash
pnpm add @office-open/odt
```

## Quick Start

```typescript
import { generateDocument, parseDocument } from "@office-open/odt";

const odt = generateDocument({
  sections: [{ children: [{ paragraph: "Hello ODT" }] }],
});
const options = parseDocument(odt);
```

## API

- `generateDocument(options, context?)` — generate an `.odt` package from `DocumentOptions`
- `parseDocument(data, options?)` — parse an `.odt` package into `DocumentOptions`
- `OdtParseError` — structured parse failure with part, path, and reason diagnostics

## Errors and Limitations

- Invalid ZIP containers, unexpected mimetypes, truncated XML, and unknown content-model elements throw `OdtParseError`.
- Generated output targets ODF 1.3.
- ODT has no patch API; parse, mutate `DocumentOptions`, and generate instead.

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
