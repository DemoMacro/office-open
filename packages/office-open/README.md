# office-open

![npm version](https://img.shields.io/npm/v/office-open)
![npm downloads](https://img.shields.io/npm/dw/office-open)
![npm license](https://img.shields.io/npm/l/office-open)

> Everything for AI-native Office documents in one install — Word (.docx), Excel (.xlsx), and PowerPoint (.pptx) generation from JSON, plus a CLI, Vercel AI SDK tools, and frozen JSON Schemas for LLM tool-calling.

## Features

- **One Install** — Import from `office-open/docx`, `office-open/pptx`, `office-open/xlsx`; no Microsoft Office required
- **AI SDK Tools** — Vercel AI SDK compatible tools for `generate-docx`, `generate-pptx`, `generate-xlsx`, ready for AI agents and chatbots
- **JSON Schemas** — Draft-07 input validation for all document types, with on-demand schema slicing for LLM context budgets
- **CLI** — Generate files from JSON via `npx office-open`
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

# Consult the JSON schemas (for AI agents and humans)
npx office-open schema index docx                        # indexed lookup entries by domain
npx office-open schema index docx --all                   # every definition name
npx office-open schema slice docx ParagraphOptions RunOptions   # sub-schema for those types
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

The generate tools carry skeleton input schemas (top-level shape + wrapper keys, ~5K tokens instead of the ~170K-token full schema); the `office-open-schema-lookup` tool fetches precise option schemas on demand.

Two more agent entry points live on the documentation site:

```bash
# MCP server for Claude Code, Cursor, and other MCP clients
claude mcp add --transport http office-open https://www.office-open.com/mcp

# Installable Agent Skill with curated API references
npx skills add https://www.office-open.com
```

### JSON Schemas

```typescript
import { validateDocumentInput, sliceDocumentSchema } from "office-open/schemas";

try {
  const validated = validateDocumentInput("docx", userInput);
} catch (e) {
  // Aggregated validation errors with instance paths and messages
}

// Extract the dependency closure of specific option types
const slice = sliceDocumentSchema("docx", ["ParagraphOptions", "RunOptions"]);
```

### Import from Sub-Packages

```typescript
import { generateDocument, parseDocument, patchDocument } from "office-open/docx";
import { generatePresentation, parsePresentation, patchPresentation } from "office-open/pptx";
import { generateWorkbook, parseWorkbook, patchWorkbook } from "office-open/xlsx";
import { convertInchesToTwip } from "office-open/core";
import { parse, stringify } from "office-open/xml";
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
| `office-open/schemas`  | JSON schemas, validation, and slicing    |

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
