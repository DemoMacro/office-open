# @office-open/odp

![npm version](https://img.shields.io/npm/v/@office-open/odp)
![npm downloads](https://img.shields.io/npm/dw/@office-open/odp)
![npm license](https://img.shields.io/npm/l/@office-open/odp)

> Generate and parse OpenDocument Presentation (`.odp`) files from plain `PresentationOptions` JSON — no Microsoft Office or LibreOffice required.

## Features

- 📈 **Canonical Presentation Model** — Read and write the same `PresentationOptions` used by PPTX and PPT
- 🧱 **Slides & Shapes** — Slide layouts, shapes, text bodies, geometry, and transforms
- 🎨 **Styles** — Graphic, text, paragraph, and page styles
- 🖼️ **Media & Charts** — Embedded images and typed charts
- 🧩 **Structured Errors** — Part, path, element, and reason diagnostics

## Installation

```bash
# pnpm
pnpm add @office-open/odp

# npm
npm install @office-open/odp

# yarn
yarn add @office-open/odp

# bun
bun add @office-open/odp
```

## Quick Start

```typescript
import { generatePresentation, parsePresentation } from "@office-open/odp";

const odp = await generatePresentation({
  slides: [{ children: [] }],
});
const options = await parsePresentation(odp);
```

## API

- `generatePresentation(options, output?)` — generate an `.odp` package from `PresentationOptions` as `Promise` (`Uint8Array` by default); `generatePresentationSync` covers synchronous output and `generatePresentationStream` returns a `ReadableStream<Uint8Array>`
- `parsePresentation(data)` — parse bytes, Base64/data-URL strings, `Blob`, or `ReadableStream` into `PresentationOptions` as `Promise`; `parsePresentationSync` covers synchronous parsing (rejects `Blob` and `ReadableStream` inputs)
- `OdpParseError` — structured parse failure with part, path, and reason diagnostics

Every input is a plain JSON object (`PresentationOptions` and its option types). The full typed API reference lives in the [documentation](https://www.office-open.com/en/getting-started/more-formats).

## Errors and Limitations

- Invalid ZIP containers, unexpected mimetypes, truncated XML, and unknown content-model elements throw `OdpParseError`.
- Generated output targets ODF 1.3.
- ODP has no patch API; parse, mutate `PresentationOptions`, and generate instead.

## Documentation

- [Documentation](https://www.office-open.com/en/getting-started/more-formats) — additional format guides and examples
- [Report Issues](https://github.com/DemoMacro/office-open/issues) — bug reports and feature requests

## Related Packages

- [office-open](https://www.npmjs.com/package/office-open) — all formats + CLI + AI SDK tools in one install
- [@office-open/pptx](https://www.npmjs.com/package/@office-open/pptx) — PowerPoint (.pptx) generation and parsing
- [@office-open/ppt](https://www.npmjs.com/package/@office-open/ppt) — legacy PowerPoint (.ppt) parsing

## License

- [MIT](https://github.com/DemoMacro/office-open/blob/main/LICENSE) © [Demo Macro](https://www.demomacro.com/)
