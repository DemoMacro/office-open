---
name: office-open
description: >
  Generate, parse, and patch Office Open XML documents (.docx, .pptx, .xlsx) using pure JSON
  options and functional API. Use when creating Word documents, PowerPoint presentations, or
  Excel spreadsheets with TypeScript — paragraphs, tables, images, charts, shapes, math equations,
  headers, footers, styles, effects, animations, data validation, or exporting to buffer/blob/base64.
  Also use when parsing existing OOXML files or patching template placeholders.
  Do not use for PDF generation, CSV files, or non-Office document formats.
---

# Office Open XML

Pure JSON + functional API for Office documents. No classes — pass option objects to `generateDocument()`, `generatePresentation()`, or `generateWorkbook()`.

## Installation

```bash
pnpm add @office-open/docx   # Word
pnpm add @office-open/pptx   # PowerPoint
pnpm add @office-open/xlsx   # Excel
pnpm add office-open          # Umbrella
```

## Workflow

1. Determine the target format: `.docx`, `.pptx`, or `.xlsx`.
2. Read the corresponding reference for the full API: `references/docx-api.md`, `references/pptx-api.md`, or `references/xlsx-api.md`.
3. Build the options object (e.g. `DocumentOptions`, `PresentationOptions`, `WorkbookOptions`).
4. Call the generate function with optional output type.

For measurement conversions (twips, EMU, half-points), see `references/units.md`.

## Generate

```ts
// DOCX
import { generateDocument } from "@office-open/docx";
const buffer = await generateDocument({ sections: [...] });
const blob = await generateDocument({ sections: [...] }, { type: "blob" });

// PPTX
import { generatePresentation } from "@office-open/pptx";
const buffer = await generatePresentation({ slides: [...] });

// XLSX
import { generateWorkbook } from "@office-open/xlsx";
const buffer = await generateWorkbook({ worksheets: [...] });
```

Output types: `"nodebuffer"` (default), `"blob"`, `"base64"`, `"uint8array"`, `"arraybuffer"`, `"string"`.
Sync variants: `generateDocumentSync`, `generatePresentationSync`, `generateWorkbookSync`.
Stream variant: `generateDocumentStream`, `generatePresentationStream`, `generateWorkbookStream`.

## Parse

Read existing files back into options objects for modification or round-trip:

```ts
import { parseDocument, generateDocument } from "@office-open/docx";
const opts = parseDocument(readFileSync("input.docx")); // Uint8Array, ArrayBuffer, base64 string
const buffer = await generateDocument(opts);
```

Same pattern: `parsePresentation` / `generatePresentation`, `parseWorkbook` / `generateWorkbook`.

## Patch

Replace `{{placeholder}}` tokens in template files:

```ts
// DOCX — paragraph or block-level replacement
import { patchDocument } from "@office-open/docx";
const result = await patchDocument({
  outputType: "nodebuffer",
  data: readFileSync("template.docx"),
  placeholders: {
    name: { type: "paragraph", children: [{ text: "John Doe" }] },
    content: { type: "document", children: [{ paragraph: { children: ["Body"] } }] },
  },
  keepOriginalStyles: true,
});

// PPTX — run-level replacement
import { patchPresentation } from "@office-open/pptx";
const result = await patchPresentation({
  outputType: "nodebuffer",
  data: readFileSync("template.pptx"),
  placeholders: { title: [{ text: "Updated" }] },
});

// XLSX — cell-level replacement
import { patchWorkbook } from "@office-open/xlsx";
const result = await patchWorkbook({
  outputType: "nodebuffer",
  data: readFileSync("template.xlsx"),
  placeholders: { amount: 1500 },
});
```

Custom delimiters: `placeholderDelimiters: { start: "{{", end: "}}" }`.

## Feature Reference

For detailed options, property tables, and complete examples, read the relevant reference file:

- **`references/docx-api.md`** — sections, text formatting, images, tables, headers/footers, math equations, styles, comments, shapes, links, patching
- **`references/pptx-api.md`** — slides, shapes, fills, outlines, effects (shadow/glow/reflection/3D), tables, images, charts, animations, transitions, groups, lines, connectors, patching
- **`references/xlsx-api.md`** — worksheets, cells, styles (font/fill/border/alignment), number formats, merge cells, freeze panes, auto filter, images, charts, data validation, conditional formatting, patching
- **`references/units.md`** — twips, EMU, half-points, 60000ths°, converter functions

## Quick Examples

### DOCX — Basic Document

```json
{
  "sections": [
    {
      "properties": {},
      "children": [
        { "paragraph": { "text": "Title", "heading": "Heading1" } },
        {
          "paragraph": {
            "children": [
              { "text": "Bold", "bold": true },
              " and ",
              { "text": "italic", "italic": true }
            ]
          }
        }
      ]
    }
  ]
}
```

### PPTX — Slide with Shapes

```json
{
  "slides": [
    {
      "children": [
        {
          "shape": {
            "x": 50,
            "y": 30,
            "width": 600,
            "height": 60,
            "textBody": { "text": "Title" },
            "fill": "4472C4"
          }
        }
      ]
    }
  ]
}
```

### XLSX — Worksheet with Data

```json
{
  "worksheets": [
    {
      "rows": [
        { "cells": [{ "value": "Name" }, { "value": "Score" }] },
        { "cells": [{ "value": "Alice" }, { "value": 95 }] }
      ]
    }
  ]
}
```

## Gotchas

- **String shorthand for text** — a plain string `"text"` in a children array is equivalent to `{ text: "text" }`.
- **Color format** — always hex without `#` prefix: `"FF0000"`, not `"#FF0000"`.
- **Font sizes** — Font sizes are in points across all formats. DOCX internally converts to half-points. PPTX and XLSX use points directly.
- **Position units** — PPTX shape positions (`x`, `y`, `width`, `height`) are in pixels, internally converted to EMU. DOCX uses twips for margins/spacing.
- **Date values in XLSX** — `Date` objects are serialized as serial numbers (days since 1899-12-30), not formatted strings. Apply `numFmt` for display formatting.
- **Charts** — chart types are defined in `@office-open/core` (`ChartSpaceOptions`), shared across docx and pptx.
- **XLSX patching** — uses `{ value: "new value" }` (cell-level), not `{ type: "paragraph", children: [...] }` like docx/pptx.
