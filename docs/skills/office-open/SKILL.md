---
name: office-open
description: >
  Generate, parse, and patch Office files with pure JSON options and a functional API. Use when
  creating or modifying Word, PowerPoint, or Excel files; parsing OOXML, legacy Office, RTF, or
  ODF; converting tabular data; or applying ordered patch commands. Do not use for PDF generation
  or unrelated document formats.
---

# Office Documents

Pure JSON + functional API for Office documents. No classes — pass option objects to `generateDocument()`, `generatePresentation()`, or `generateWorkbook()`.

## Installation

```bash
pnpm add @office-open/docx   # Word
pnpm add @office-open/pptx   # PowerPoint
pnpm add @office-open/xlsx   # Excel
pnpm add office-open          # Umbrella
```

## Workflow

1. Route the request by format family and capability:

   | Request                                          | API                                                  |
   | ------------------------------------------------ | ---------------------------------------------------- |
   | Create/edit document/workbook/presentation files | Family dispatch or format package generate/patch API |
   | Read `.doc`, `.xls`, `.ppt`                      | `parseOffice`, then use the projected OOXML model    |
   | Read/write `.rtf`, `.odt`, `.ods`, `.odp`        | Family or `office-open` dispatch API                 |
   | Convert CSV/TSV                                  | `office-open/convert`                                |

2. Read the corresponding reference for the full API: `references/docx-api.md`, `references/pptx-api.md`, or `references/xlsx-api.md`. For the family dispatch and capability matrix, see the umbrella package README.
3. Build the options object (e.g. `DocumentOptions`, `PresentationOptions`, `WorkbookOptions`).
4. Call the generate function with optional output type.
5. Verify the result by parsing it back and checking the canonical fields that the request changed; OOXML patch APIs preserve unrelated package facts, while RTF/ODF patch APIs preserve canonical and package-overlay fields.

For measurement conversions (twips, EMU, half-points), see `references/units.md`. For exact field names, enum values, or unit constraints of any options type, see [Precise Field Lookup](#precise-field-lookup) below.

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

Output types: `"uint8array"` (default), `"blob"`, `"base64"`, `"arraybuffer"`, `"string"`, `"nodebuffer"`.
Sync variants: `generateDocumentSync`, `generatePresentationSync`, `generateWorkbookSync`.
Stream variant: `generateDocumentStream`, `generatePresentationStream`, `generateWorkbookStream`.

## Parse

Read existing files back into options objects for modification or round-trip:

```ts
import { parseDocument, generateDocument } from "@office-open/docx";
const opts = parseDocument(readFileSync("input.docx")); // Uint8Array, ArrayBuffer, typed-array view, base64 string
const buffer = await generateDocument(opts);
```

Same pattern: `parsePresentation` / `generatePresentation`, `parseWorkbook` / `generateWorkbook`.

## Patch

Patch template files with ordered, source-preserving commands:

```ts
// DOCX — text and block content
import { patchDocument } from "@office-open/docx";
const result = await patchDocument({
  outputType: "uint8array",
  data: readFileSync("template.docx"),
  commands: [
    {
      type: "replace",
      at: { field: "sections", route: [{ index: 0 }], text: { placeholder: "name" } },
      value: "John Doe",
    },
    {
      type: "update",
      at: { field: "sections", route: [{ index: 0 }] },
      set: { children: [{ paragraph: { children: ["Body"] } }] },
    },
    {
      type: "insert",
      at: { field: "sections", route: [{ index: 0 }, { collection: "children" }] },
      position: "end",
      value: [{ paragraph: { children: ["Appended"] } }],
    },
  ],
});

// PPTX — text and slides
import { patchPresentation } from "@office-open/pptx";
const result = await patchPresentation({
  outputType: "uint8array",
  data: readFileSync("template.pptx"),
  commands: [
    {
      type: "replace",
      at: { field: "slides", route: [{ index: 0 }], text: { placeholder: "title" } },
      value: "Updated",
    },
    { type: "remove", at: { field: "slides", route: [{ index: 1 }] } },
  ],
});

// XLSX — text and worksheets
import { patchWorkbook } from "@office-open/xlsx";
const result = await patchWorkbook({
  outputType: "uint8array",
  data: readFileSync("template.xlsx"),
  commands: [
    {
      type: "replace",
      at: { field: "worksheets", route: [{ name: "Draft" }], text: { placeholder: "amount" } },
      value: 1500,
    },
    { type: "remove", at: { field: "worksheets", route: [{ name: "Draft" }] } },
  ],
});
```

Address each command with `at: { field, route, text }`: `field` is the canonical root field, `route` selects nested nodes, and optional `text` selects placeholder or literal text. Pass `value`, `set`, or `to` for `replace|insert`, `update`, and `move` respectively. Set `placeholderDelimiters: { start: "{{", end: "}}" }` to customize placeholder tokens.

Patch is available for writable OOXML, RTF, and ODF formats. RTF patching regenerates the file after canonical edits; ODF patching preserves canonical package overlays. Legacy DOC/XLS/PPT remains read-only. Never invent patch fields outside the canonical model.

## Feature Reference

For detailed options, property tables, and complete examples, read the relevant reference file:

- **`references/docx-api.md`** — sections, text formatting, images, tables, headers/footers, math equations, styles, comments, shapes, links, patching
- **`references/pptx-api.md`** — slides, shapes, fills, outlines, effects (shadow/glow/reflection/3D), tables, images, charts, animations, transitions, groups, lines, connectors, patching
- **`references/xlsx-api.md`** — worksheets, cells, styles (font/fill/border/alignment), number formats, merge cells, freeze panes, auto filter, images, charts, data validation, conditional formatting, patching
- **`references/units.md`** — twips, EMU, half-points, 60000ths°, converter functions

## Precise Field Lookup

The references above are curated summaries. When you need the exact fields of any options type (enum values, unit constraints, optionality), query the package's JSON Schema instead of guessing:

```bash
npx office-open schema index docx                          # indexed lookup entries by domain
npx office-open schema index docx --all                   # every definition name
npx office-open schema slice docx ParagraphOptions RunOptions   # sub-schema for those types
```

A slice carries the full dependency closure of the requested types; other cataloged domains stay as expandable stubs. Generic definition names need shell quotes: `"HeaderFooterGroup<HeaderFooterReference>"`. TypeScript remains the source of truth at compile time — use the CLI to read, not to validate, your code.

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
- **Geometry units** — DrawingML geometry accepts EMU numbers or `UniversalMeasure` strings (`"5cm"`, `"2in"`, `"200px"`). DOCX uses twips for margins/spacing.
- **Date values in XLSX** — `Date` objects are serialized as serial numbers (days since 1899-12-30), not formatted strings. Apply `numFmt` for display formatting.
- **Charts** — chart types are defined in `@office-open/core` (`ChartSpaceOptions`), shared across docx and pptx.
- **XLSX patching** — text `replace` accepts scalar values; whole worksheets are addressed through `worksheets` routes with `WorksheetOptions`.
