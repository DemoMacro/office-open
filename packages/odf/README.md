# @office-open/odf

![npm version](https://img.shields.io/npm/v/@office-open/odf)
![npm downloads](https://img.shields.io/npm/dw/@office-open/odf)
![npm license](https://img.shields.io/npm/l/@office-open/odf)

> Generate and parse OpenDocument `.odt`, `.ods`, and `.odp` files through the office-open Options models — no LibreOffice required.

## Features

- 📄 **Text Documents** — Generate and parse `.odt` paragraphs, headings, rich text, and tables
- 📊 **Spreadsheets** — Generate and parse `.ods` worksheets, cells, formulas, columns, rows, and visibility
- 📽️ **Presentations** — Generate and parse `.odp` slides, shapes, rich text, and page geometry
- 🎨 **Basic Styles** — Preserve common text, paragraph, dimension, and visibility styles
- 📝 **Metadata** — Round-trip titles, subjects, creators, keywords, dates, categories, and revisions
- 🧩 **Generic ODF Nodes** — Address ODF elements outside the specialized typed surface as extension nodes
- 📦 **Low-Level Package APIs** — Read OCF manifests, package files, XML nodes, charts, and database documents

The specialized ODF surface is intentionally compact today. Generic node support keeps all ODF Relax NG elements addressable for parsing and serialization; specialized typed coverage continues to expand separately.

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
import { generateOdt, parseOdt } from "@office-open/odf";
import { readFileSync, writeFileSync } from "node:fs";

const options = {
  sections: [{ children: [{ paragraph: "Hello ODF" }] }],
};
const buffer = generateOdt(options);
writeFileSync("output.odt", buffer);

const parsed = parseOdt(new Uint8Array(readFileSync("output.odt")));
```

Spreadsheets and presentations use the same parse/generate pairing:

```typescript
import { generateOds, generateOdp, parseOds, parseOdp } from "@office-open/odf";

const spreadsheet = generateOds(spreadsheetOptions);
const presentation = generateOdp(presentationOptions);
const parsedSpreadsheet = parseOds(spreadsheet);
const parsedPresentation = parseOdp(presentation);
```

## API

- `generateOdt(options)` / `parseOdt(data)` — ODF Text (`OdtOptions`)
- `generateOds(options)` / `parseOds(data)` — ODF Spreadsheet (`OdsOptions`)
- `generateOdp(options)` / `parseOdp(data)` — ODF Presentation (`OdpOptions`)
- `generateChartDocument(options)` / `parseChartDocument(data)` — ODF Chart (`ChartDocumentOptions`)
- `generateDatabaseDocument(options)` / `parseDatabaseDocument(data)` — ODF Database (`DatabaseDocumentOptions`)
- `parseOdfNode`, `parseOdfNodes`, `serializeOdfNodes`, `ODF_ELEMENT_NAMES` — generic ODF node extensions
- `generateOcf`, `manifestXml`, `readOcf`, `readXml` — ODF container and manifest helpers

The high-level models follow the corresponding office-open `DocumentOptions`, `WorkbookOptions`, and `PresentationOptions` shapes with ODF-specific extensions. ODF packages use an uncompressed leading `mimetype` and an ODF 1.3 manifest.

## Documentation

- [Documentation](https://www.office-open.com/en/getting-started/more-formats) — additional format guides and examples
- [Report Issues](https://github.com/DemoMacro/office-open/issues) — bug reports and feature requests

## Related Packages

- [office-open](https://www.npmjs.com/package/office-open) — all formats + CLI + AI SDK tools in one install
- [@office-open/docx](https://www.npmjs.com/package/@office-open/docx) — Word (.docx) generation and parsing
- [@office-open/xlsx](https://www.npmjs.com/package/@office-open/xlsx) — Excel (.xlsx) generation and parsing
- [@office-open/pptx](https://www.npmjs.com/package/@office-open/pptx) — PowerPoint (.pptx) generation and parsing

## License

- [MIT](https://github.com/DemoMacro/office-open/blob/main/LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
