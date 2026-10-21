# @office-open/odf

![npm version](https://img.shields.io/npm/v/@office-open/odf)
![npm downloads](https://img.shields.io/npm/dw/@office-open/odf)
![npm license](https://img.shields.io/npm/l/@office-open/odf)

> Read and write OpenDocument (OCF) packages — the shared runtime powering ODT, ODS, and ODP.

## Features

- 📦 **OCF Packages** — Read and generate ZIP containers, `mimetype`, and manifests
- 🧩 **Typed Runtime** — Parse ODF XML into a generic runtime model with structured errors
- 📊 **Shared Codecs** — Chart, database, form, drawing, metadata, and manifest models
- 🔍 **Semantic Digests** — Compare OCF entries, manifest data, and binary content

`@office-open/odf` is the implementation runtime for `@office-open/odt`, `@office-open/ods`, and `@office-open/odp`. Use those packages for canonical document models.

## Installation

```bash
# pnpm
pnpm add @office-open/odf

# npm
npm install @office-open/odf

# yarn
yarn add @office-open/odf

# bun
bun add @office-open/odf
```

## Quick Start

```typescript
import { readOcf } from "@office-open/odf";

const archive = await readOcf(file, "application/vnd.oasis.opendocument.text");
```

## API

- `readOcf(data, expectedMimeType)` — inspect and parse bytes, Base64/data-URL strings, `Blob`, or `ReadableStream` as an OCF package
- `generateOcf(mimeType, files, mediaTypes?, packageManifest?)` — generate an OCF package
- `parseManifestOptions(xml)` / `manifestOptionsXml(options)` — map manifest XML
- `parseMeta(xml)` / `metaXml(options)` — map ODF metadata to core properties
- `OcfError` and `OdfSchemaError` — structured container and schema failures

## Documentation

- [Documentation](https://www.office-open.com/en/getting-started/more-formats) — OpenDocument guides and examples
- [Report Issues](https://github.com/DemoMacro/office-open/issues) — bug reports and feature requests

## Related Packages

- [office-open](https://www.npmjs.com/package/office-open) — all formats + CLI + AI SDK tools in one install
- [@office-open/odt](https://www.npmjs.com/package/@office-open/odt) — OpenDocument Text
- [@office-open/ods](https://www.npmjs.com/package/@office-open/ods) — OpenDocument Spreadsheet
- [@office-open/odp](https://www.npmjs.com/package/@office-open/odp) — OpenDocument Presentation

## License

- [MIT](https://github.com/DemoMacro/office-open/blob/main/LICENSE) © [Demo Macro](https://www.demomacro.com/)
