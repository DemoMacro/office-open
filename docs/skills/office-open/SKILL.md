---
name: office-open
description: >
  Generate and parse Office Open XML documents (.docx, .pptx, .xlsx) with JavaScript/TypeScript.
  Use when creating Word documents, PowerPoint presentations, or Excel spreadsheets, working with paragraphs,
  tables, images, charts, shapes, math equations, headers, footers, styles, themes,
  effects, animations, or exporting to buffer/blob/base64.
---

# Office Open XML

Generate and parse OOXML documents (.docx, .pptx, .xlsx) with a declarative TypeScript API. Works in Node.js and browsers.

## Installation

```bash
pnpm add @office-open/docx      # Word documents
pnpm add @office-open/pptx      # PowerPoint presentations
pnpm add @office-open/xlsx      # Excel spreadsheets
pnpm add @office-open/core      # Shared utilities (usually not needed directly)
pnpm add @office-open/xml       # XML parsing/serialization
pnpm add office-open             # Umbrella: all packages + CLI + AI tools
```

## Packages

| Package             | Purpose                                 |
| ------------------- | --------------------------------------- |
| `@office-open/docx` | Word document generation & parsing      |
| `@office-open/pptx` | PowerPoint generation & parsing         |
| `@office-open/xlsx` | Spreadsheet generation & parsing        |
| `@office-open/core` | Shared XML components, converters       |
| `@office-open/xml`  | Low-level XML parse/stringify/query     |
| `office-open`       | Umbrella: all packages + CLI + AI tools |

## Quick Start — DOCX

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
              { "text": "italic", "italics": true }
            ]
          }
        }
      ]
    }
  ]
}
```

```ts
import { Document, Packer } from "@office-open/docx";

// The JSON above is the Document options — pass to constructor
const doc = new Document({ sections: [...] });
const buffer = await Packer.toBuffer(doc);
```

## Quick Start — PPTX

```json
{
  "title": "Demo",
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

```ts
import { Presentation, Packer } from "@office-open/pptx";

// Pass options to Presentation constructor
const pres = new Presentation({ title: "Demo", slides: [...] });
const buffer = await Packer.toBuffer(pres);
```

## Quick Start — XLSX

```json
{
  "worksheets": [
    {
      "children": [
        { "cells": [{ "value": "Name" }, { "value": "Score" }] },
        { "cells": [{ "value": "Alice" }, { "value": 95 }] }
      ]
    }
  ]
}
```

```ts
import { Workbook, Packer } from "@office-open/xlsx";

// Pass options to Workbook constructor
const wb = new Workbook({ worksheets: [...] });
const buffer = await Packer.toBuffer(wb);
```

## Export Options

`Packer` provides multiple output formats:

| Method                | Returns                      | Environment |
| --------------------- | ---------------------------- | ----------- |
| `toBuffer(doc)`       | `Promise<Buffer>`            | Node.js     |
| `toBlob(doc)`         | `Promise<Blob>`              | Browser     |
| `toBase64String(doc)` | `Promise<string>`            | Any         |
| `toBytes(doc)`        | `Promise<Uint8Array>`        | Any         |
| `toArrayBuffer(doc)`  | `Promise<ArrayBuffer>`       | Any         |
| `toString(doc)`       | `Promise<string>`            | Any         |
| `toStream(doc)`       | `ReadableStream<Uint8Array>` | Any         |

Each async method has a sync counterpart: `toBufferSync`, `toBlobSync`, `toBase64StringSync`, etc.

## Parsing Existing Documents

```ts
import { parseDocument, Document, Packer } from "@office-open/docx";
import { readFileSync } from "node:fs";

// Accepts Uint8Array, ArrayBuffer, DataView, number[], base64 string
const opts = parseDocument(readFileSync("input.docx"));
const doc = new Document(opts);
const buffer = await Packer.toBuffer(doc);
```

```ts
import { parsePresentation, Presentation } from "@office-open/pptx";

const presOpts = parsePresentation(readFileSync("input.pptx"));
const pres = new Presentation(presOpts);
const presBuffer = await Packer.toBuffer(pres);
```

```ts
import { parseWorkbook, Workbook, Packer } from "@office-open/xlsx";

const wbOpts = parseWorkbook(readFileSync("input.xlsx"));
const wb = new Workbook(wbOpts);
const wbBuffer = await Packer.toBuffer(wb);
```

## Patching Existing Documents

Replace `{{placeholder}}` tokens in existing templates with new content.

### DOCX

```ts
import { patchDocument, PatchType, TextRun, Paragraph, Table } from "@office-open/docx";

// Replace paragraph-level placeholder
const result = await patchDocument({
  outputType: "nodebuffer",
  data: readFileSync("template.docx"),
  patches: {
    name: { type: PatchType.PARAGRAPH, children: [new TextRun("John Doe")] },
    content: {
      type: PatchType.DOCUMENT,
      children: [new Paragraph("First"), new Table({ rows: [...] })],
    },
  },
  keepOriginalStyles: true,
  recursive: true,
});
```

### PPTX

```ts
import { patchPresentation, PatchType, TextRun } from "@office-open/pptx";

const result = await patchPresentation({
  outputType: "nodebuffer",
  data: readFileSync("template.pptx"),
  patches: {
    title: { type: PatchType.PARAGRAPH, children: [new TextRun({ text: "Updated", bold: true })] },
  },
  keepOriginalStyles: true,
});
```

### XLSX

```ts
import { patchWorkbook, PatchType, TextRun } from "@office-open/xlsx";

const result = await patchWorkbook({
  outputType: "nodebuffer",
  data: readFileSync("template.xlsx"),
  patches: {
    name: { type: PatchType.PARAGRAPH, children: [new TextRun("John Doe")] },
  },
  keepOriginalStyles: true,
});
```

Both functions support custom `placeholderDelimiters` (default: `{{` and `}}`). DOCX also provides `patchDetector` to scan templates for placeholder keys.

## DOCX Features

### Sections & Page Layout

```json
{
  "sections": [
    {
      "properties": {
        "page": {
          "size": { "width": 11906, "height": 16838, "orientation": "portrait" },
          "margin": { "top": 1440, "bottom": 1440, "left": 1440, "right": 1440 }
        }
      },
      "children": []
    }
  ]
}
```

### Tables

```json
{
  "width": { "size": 100, "type": "pct" },
  "rows": [
    {
      "children": [
        { "children": [{ "paragraph": { "text": "Cell 1" } }], "shading": { "fill": "4472C4" } },
        { "children": [{ "paragraph": { "text": "Cell 2" } }] }
      ]
    }
  ]
}
```

### Images with Effects

```json
{
  "type": "png",
  "data": "<Uint8Array>",
  "transformation": { "width": 200, "height": 150 },
  "blipEffects": {
    "grayscale": true,
    "luminance": { "bright": 30, "contrast": -20 },
    "duotone": { "color1": { "value": "002060" }, "color2": { "value": "D0CECE" } },
    "tint": { "hue": 6000000, "amount": 40 }
  }
}
```

### Headers & Footers

```json
{
  "headers": {
    "default": {
      "children": [{ "paragraph": { "text": "Header text" } }]
    }
  },
  "footers": {
    "default": {
      "children": [
        {
          "paragraph": {
            "alignment": "center",
            "children": ["Page ", "CURRENT", " of ", "TOTAL_PAGES"]
          }
        }
      ]
    }
  }
}
```

Page number tokens (string values used in TextRun children arrays):

- `"CURRENT"` — Current page number
- `"TOTAL_PAGES"` — Total page count
- `"TOTAL_PAGES_IN_SECTION"` — Section page count
- `"SECTION"` — Current section page

### Math Equations

```json
// Fraction: a/b
{ "fraction": { "numerator": ["a"], "denominator": ["b"] } }

// Superscript: x²
{ "superScript": { "children": ["x"], "superScript": ["2"] } }

// Square root: √(a+b)
{ "radical": { "children": ["a + b"] } }
```

### Charts & SmartArt

Charts are created via `@office-open/core` and embedded in both docx and pptx.

```json
{
  "type": "bar",
  "categories": ["A", "B"],
  "series": [{ "name": "Series1", "values": [10, 20] }]
}
```

## PPTX Features

### Shapes with Text

```json
{
  "shape": {
    "x": 50,
    "y": 100,
    "width": 300,
    "height": 80,
    "textBody": { "text": "Hello" },
    "fill": "4472C4",
    "outline": { "color": "0D47A1", "width": 12700 },
    "geometry": "rect"
  }
}
```

### Shapes with Effects

```json
{
  "shape": {
    "x": 50,
    "y": 100,
    "width": 200,
    "height": 120,
    "textBody": { "text": "Shadow" },
    "fill": "ED7D31",
    "effects": {
      "outerShadow": {
        "blur": 50800,
        "distance": 38100,
        "direction": 5400000,
        "color": "000000",
        "alpha": 50
      },
      "glow": { "radius": 152400, "color": "92D050", "alpha": 60 },
      "reflection": { "blurRadius": 6350, "distance": 38100, "startAlpha": 90, "endAlpha": 0 },
      "softEdge": { "radius": 50800 },
      "rotation3D": { "x": 20, "y": 30, "z": 10, "perspective": 500 },
      "extrusionH": 50000,
      "material": "plastic",
      "bevelTop": { "width": 8, "height": 8 }
    }
  }
}
```

### Fill Types

```json
{ "type": "solid", "color": "4472C4" }
{ "type": "gradient", "stops": [{ "position": 0, "color": "1565C0" }, { "position": 100, "color": "42A5F5" }], "angle": 45 }
{ "type": "blip", "data": "<Uint8Array>", "imageType": "png" }
{ "type": "none" }
```

String shorthand: `fill: "4472C4"` is equivalent to `{ type: "solid", color: "4472C4" }`.

## Measurement Units

| Unit        | Used For                  | Conversion          |
| ----------- | ------------------------- | ------------------- |
| Twips       | Margins, spacing          | 1 inch = 1440 twips |
| EMU         | Images, shapes, offsets   | 1 inch = 914400 EMU |
| Half-points | Font sizes                | Size 24 = 12pt      |
| 60000ths°   | Angles, rotation, effects | 45° = 2700000       |

## References

See the `references/` directory for detailed API documentation on each topic.
