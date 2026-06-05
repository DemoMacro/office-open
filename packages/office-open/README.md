# office-open

![npm version](https://img.shields.io/npm/v/office-open)
![npm downloads](https://img.shields.io/npm/dw/office-open)
![npm license](https://img.shields.io/npm/l/office-open)

> Umbrella package for Office Open XML — all packages, CLI, AI SDK tools, and Zod schemas in one install.

## Features

- **Unified Re-exports** — Import from `office-open/docx`, `office-open/pptx`, `office-open/xlsx`
- **CLI** — Generate files from JSON via `npx office-open`
- **AI SDK Tools** — Vercel AI SDK compatible tools for `generate-docx`, `generate-pptx`, `generate-xlsx`
- **Zod Schemas** — Input validation for all document types
- **Generate Function** — Type-agnostic `generate()` for dynamic document creation

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

### Generate from JSON

```typescript
import { generate, generateToFile } from "office-open/generate";

const buffer = await generate({
  type: "docx",
  options: {
    sections: [
      {
        children: [{ paragraph: "Hello World" }],
      },
    ],
  },
  outputType: "nodebuffer",
});
```

### CLI

```bash
# Generate from a JSON file
npx office-open docx document.json "output.docx"
npx office-open pptx slides.json "output.pptx"
npx office-open xlsx spreadsheet.json "output.xlsx"
```

### AI SDK Tools

```typescript
import { officeOpenTools } from "office-open/ai";

// Use with Vercel AI SDK
const result = await generateText({
  model,
  tools: officeOpenTools,
  prompt: "Create a sales report as a .docx file",
});
```

### Zod Schemas

```typescript
import { validateDocumentInput } from "office-open/schemas";

try {
  const validated = validateDocumentInput("docx", userInput);
} catch (e) {
  // Structured validation error with path and message
}
```

### Import from Sub-Packages

```typescript
import { Document, Packer } from "office-open/docx";
import { Presentation, Packer } from "office-open/pptx";
import { Workbook, Packer } from "office-open/xlsx";
import { convertInchesToTwip } from "office-open/core";
import { xml2js, js2xml } from "office-open/xml";
```

## Sub-Exports

| Export Path            | Description                              |
| ---------------------- | ---------------------------------------- |
| `office-open`          | Main entry (re-exports all sub-packages) |
| `office-open/docx`     | @office-open/docx                        |
| `office-open/pptx`     | @office-open/pptx                        |
| `office-open/xlsx`     | @office-open/xlsx                        |
| `office-open/core`     | @office-open/core                        |
| `office-open/xml`      | @office-open/xml                         |
| `office-open/generate` | `generate()` function                    |
| `office-open/ai`       | Vercel AI SDK tools                      |
| `office-open/schemas`  | Zod validation schemas                   |

## JSON Document Structures

### DOCX

```json
{
  "sections": [{ "children": [{ "paragraph": "Hello World" }] }]
}
```

### PPTX

```json
{
  "title": "My Deck",
  "slides": [{ "children": [{ "shape": { "textBody": { "text": "Hello" } } }] }]
}
```

### XLSX

```json
{
  "worksheets": [{ "rows": [{ "cells": [{ "value": "Name" }, { "value": 95 }] }] }]
}
```

## License

MIT
