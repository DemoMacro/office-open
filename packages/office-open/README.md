# office-open

![npm version](https://img.shields.io/npm/v/office-open)
![npm downloads](https://img.shields.io/npm/dw/office-open)
![npm license](https://img.shields.io/npm/l/office-open)

> Generate, parse, and patch AI-native Office documents with three canonical option models — plus a CLI, AI SDK tools, JSON Schemas, CSV/TSV conversion, and support for OOXML, legacy, RTF, and OpenDocument formats.

## Features

- 🧠 **Canonical Models** — `DocumentOptions`, `WorkbookOptions`, and `PresentationOptions`
- 📄 **Word** — DOCX generation, parsing, and patching; DOC parsing
- 📊 **Excel** — XLSX generation, parsing, and patching; XLS parsing
- 📈 **PowerPoint** — PPTX generation, parsing, and patching; PPT parsing
- 📝 **RTF** — Strict generation and parsing through `DocumentOptions`
- 📦 **OpenDocument** — ODT, ODS, and ODP generation and parsing
- 🔁 **Cross-Format** — Shared drawing, chart, picture, table, connector, and group concepts
- 🤖 **AI Tools** — Vercel AI SDK tools with schema validation and structured errors
- 🧩 **JSON Schemas** — Draft-07 schemas and dependency-closed schema slicing
- 🛠️ **CLI** — Detect, parse, generate, and inspect schemas from the terminal
- 🔄 **CSV/TSV** — Convert tabular data to and from `WorkbookOptions`

## Installation

```bash
# pnpm
pnpm add office-open

# npm
npm install office-open

# yarn
yarn add office-open

# bun
bun add office-open
```

## Quick Start

### Generate

```typescript
import { generateOffice } from "office-open";
import { writeFileSync } from "node:fs";

const buffer = await generateOffice("docx", {
  sections: [{ children: [{ paragraph: "Hello World" }] }],
});

writeFileSync("Hello.docx", buffer);
```

### Parse

```typescript
import { detectOffice, parseOffice } from "office-open";
import { readFileSync } from "node:fs";

const input = new Uint8Array(readFileSync("input.docx"));
console.log(detectOffice(input).format);

const parsed = await parseOffice(input);
console.log(parsed.format, parsed.options);
```

### Patch

```typescript
import { patchDocument } from "office-open";

const patched = await patchDocument({
  format: "docx",
  data: input,
  findReplace: {
    "Account Manager": "Jane Doe",
  },
});
```

### Family Entries

```typescript
import { generateDocument, parseDocument, patchDocument } from "office-open/document";
import { generateWorkbook, parseWorkbook, patchWorkbook } from "office-open/workbook";
import {
  generatePresentation,
  parsePresentation,
  patchPresentation,
} from "office-open/presentation";
```

### Per-Format Entries

```typescript
import { generateDocument, parseDocument, patchDocument } from "office-open/docx";
import { generatePresentation, parsePresentation, patchPresentation } from "office-open/pptx";
import { generateWorkbook, parseWorkbook, patchWorkbook } from "office-open/xlsx";
```

### CLI

```bash
npx office-open docx document.json "output.docx"
npx office-open rtf document.json "output.rtf"
npx office-open detect input.doc
npx office-open parse input.xlsx --output workbook.json
npx office-open schema slice docx ParagraphOptions RunOptions
```

## API

- `detectOffice(input)` — identify format, family, and container from bytes or RTF text
- `parseOffice(input, options?)` — parse any supported format into a discriminated canonical model
- `generateOffice(format, options, output?)` — generate any writable format
- `parseDocument` / `generateDocument` / `patchDocument` — document family dispatch
- `parseWorkbook` / `generateWorkbook` / `patchWorkbook` — workbook family dispatch
- `parsePresentation` / `generatePresentation` / `patchPresentation` — presentation family dispatch

Additional entries include `office-open/convert`, `office-open/schemas`, `office-open/ai`, `office-open/core`, and `office-open/xml`.

## Format Support

| Family       | Formats                               | Parse |              Generate |       Patch |
| ------------ | ------------------------------------- | ----: | --------------------: | ----------: |
| Document     | DOC, DOCX, DOCM, DOTX, DOTM, RTF, ODT |    ✅ | DOCX family, RTF, ODT | DOCX family |
| Workbook     | XLS, XLSX, XLSM, XLTX, XLTM, ODS      |    ✅ |      XLSX family, ODS | XLSX family |
| Presentation | PPT, PPTX, PPTM, POTX, POTM, ODP      |    ✅ |      PPTX family, ODP | PPTX family |
| Tabular      | CSV, TSV                              |    ✅ |                    ✅ |          ❌ |

Encrypted OOXML inputs are detected precisely and returned as parse errors with actionable context rather than being parsed as unrelated formats.

## Canonical Model

There are three canonical models, not one model per file extension:

```typescript
type DocumentOptions = /* DOC, DOCX, RTF, and ODT projection */;
type WorkbookOptions = /* XLS, XLSX, and ODS projection */;
type PresentationOptions = /* PPT, PPTX, and ODP projection */;
```

Every parser returns the same family model. Therefore legacy and ODF content can flow into the same DOCX, XLSX, and PPTX generation and conversion paths without a fourth ODF document model.

## Errors and Limitations

- DOC, XLS, and PPT are parse-only; legacy binary writers are not provided.
- RTF generation rejects options that cannot be represented in RTF instead of dropping them.
- ODT, ODS, and ODP support parse and generate but do not provide a patch API.
- Parse errors expose structured context such as part, path, element name, and reason.
- Unsupported operations throw explicit errors; they do not silently change format.

## AI SDK Tools

```typescript
import { officeOpenTools } from "office-open/ai";

const result = await generateText({
  model,
  tools: officeOpenTools,
  prompt: "Create a sales report as a .docx file",
});
```

The tools use skeleton schemas by default. Use `officeOpenTools({ fullSchemas: true })` when the model can afford the complete schemas.

## JSON Schemas

```typescript
import { sliceDocumentSchema, validateDocumentInput } from "office-open/schemas";

const input = validateDocumentInput("docx", userInput);
const slice = sliceDocumentSchema("docx", ["ParagraphOptions", "RunOptions"]);
```

## Documentation

- [Documentation](https://www.office-open.com) — getting started, per-format guides, and AI integration
- [More Formats](https://www.office-open.com/en/getting-started/more-formats) — RTF and OpenDocument examples
- [Report Issues](https://github.com/DemoMacro/office-open/issues) — bug reports and feature requests

## Related Packages

- [@office-open/docx](https://www.npmjs.com/package/@office-open/docx) — Word (.docx)
- [@office-open/pptx](https://www.npmjs.com/package/@office-open/pptx) — PowerPoint (.pptx)
- [@office-open/xlsx](https://www.npmjs.com/package/@office-open/xlsx) — Excel (.xlsx)
- [@office-open/doc](https://www.npmjs.com/package/@office-open/doc) — legacy Word (.doc)
- [@office-open/xls](https://www.npmjs.com/package/@office-open/xls) — legacy Excel (.xls)
- [@office-open/ppt](https://www.npmjs.com/package/@office-open/ppt) — legacy PowerPoint (.ppt)
- [@office-open/rtf](https://www.npmjs.com/package/@office-open/rtf) — Rich Text Format
- [@office-open/odt](https://www.npmjs.com/package/@office-open/odt) — OpenDocument Text
- [@office-open/ods](https://www.npmjs.com/package/@office-open/ods) — OpenDocument Spreadsheet
- [@office-open/odp](https://www.npmjs.com/package/@office-open/odp) — OpenDocument Presentation
- [@office-open/core](https://www.npmjs.com/package/@office-open/core) — shared OOXML infrastructure
- [@office-open/xml](https://www.npmjs.com/package/@office-open/xml) — XML parsing and serialization

## License

- [MIT](https://github.com/DemoMacro/office-open/blob/main/LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
