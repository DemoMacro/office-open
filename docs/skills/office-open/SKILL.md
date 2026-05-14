---
name: office-open
description: >
  Generate and parse Office Open XML documents (.docx, .pptx) with JavaScript/TypeScript.
  Use when creating Word documents or PowerPoint presentations, working with paragraphs,
  tables, images, charts, shapes, math equations, headers, footers, styles, themes,
  effects, animations, or exporting to buffer/blob/base64.
---

# Office Open XML

Generate and parse OOXML documents (.docx, .pptx) with a declarative TypeScript API. Works in Node.js and browsers.

## Installation

```bash
pnpm add @office-open/docx      # Word documents
pnpm add @office-open/pptx      # PowerPoint presentations
pnpm add @office-open/core      # Shared utilities (usually not needed directly)
pnpm add @office-open/xml       # XML parsing/serialization
```

## Packages

| Package                | Purpose                              |
| ---------------------- | ------------------------------------ |
| `@office-open/docx`   | Word document generation & parsing   |
| `@office-open/pptx`   | PowerPoint generation & parsing      |
| `@office-open/core`   | Shared XML components, converters    |
| `@office-open/xml`    | Low-level XML parse/stringify/query  |

## Quick Start — DOCX

```ts
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from "@office-open/docx";

const doc = new Document({
    sections: [{
        properties: {},
        children: [
            new Paragraph({ text: "Title", heading: HeadingLevel.HEADING_1 }),
            new Paragraph({
                children: [
                    new TextRun({ text: "Bold", bold: true }),
                    new TextRun(" and "),
                    new TextRun({ text: "italic", italics: true }),
                ],
            }),
        ],
    }],
});

const buffer = await Packer.toBuffer(doc);
```

## Quick Start — PPTX

```ts
import { Presentation, Packer, Shape } from "@office-open/pptx";

const pres = new Presentation({
    title: "Demo",
    slides: [{
        children: [
            new Shape({
                x: 50, y: 30, width: 600, height: 60,
                text: "Title",
                fill: "4472C4",
            }),
        ],
    }],
});

const buffer = await Packer.toBuffer(pres);
```

## Export Options

```ts
// Node.js
const buffer = await Packer.toBuffer(doc);

// Browser
const blob = await Packer.toBlob(doc);

// Base64
const base64 = await Packer.toBase64(doc);

// Write to file system
await Packer.toFile(doc, "output.docx");
```

## Parsing Existing Documents

```ts
import { Packer, Document } from "@office-open/docx";
import { readFileSync } from "node:fs";

const buffer = readFileSync("input.docx");
const doc = await Packer.fromBuffer(buffer);

// Access parsed content
const parsed = doc.toParsedDocument();
```

## DOCX Features

### Sections & Page Layout

```ts
new Document({
    sections: [{
        properties: {
            page: {
                size: { width: 11906, height: 16838, orientation: "portrait" },
                margins: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
            },
        },
        children: [...],
    }],
});
```

### Tables

```ts
import { Table, TableRow, TableCell, WidthType } from "@office-open/docx";

new Table({
    rows: [
        new TableRow({
            children: [
                new TableCell({ children: [new Paragraph("Cell 1")] }),
                new TableCell({ children: [new Paragraph("Cell 2")] }),
            ],
        }),
    ],
});
```

### Images with Effects

```ts
import { ImageRun } from "@office-open/docx";
import { readFileSync } from "node:fs";

new ImageRun({
    type: "png",
    data: readFileSync("photo.png"),
    transformation: { width: 200, height: 150 },
    blipEffects: {
        grayscale: true,
        // luminance: { bright: 30, contrast: -20 },
        // duotone: { color1: { value: "002060" }, color2: { value: "D0CECE" } },
        // tint: { hue: 6000000, amount: 40 },
    },
});
```

### Headers & Footers

```ts
import { Header, Footer, PageNumber } from "@office-open/docx";

new Document({
    sections: [{
        properties: {},
        headers: {
            default: new Header({
                children: [new Paragraph("Header text")],
            }),
        },
        footers: {
            default: new Footer({
                children: [
                    new Paragraph({
                        children: [PageNumber.CURRENT],
                    }),
                ],
            }),
        },
        children: [...],
    }],
});
```

### Math Equations

```ts
import { Math, MathRun, MathFraction, MathSuperScript, MathRadical } from "@office-open/docx";

new Math({
    children: [
        new MathFraction({ numerator: "a + b", denominator: "c" }),
    ],
});

// E = mc²
new Math({
    children: [
        new MathRun("E = m"),
        new MathSuperScript({
            children: [new MathRun("c")],
            superScript: [new MathRun("2")],
        }),
    ],
});
```

### Charts & SmartArt

```ts
import { ChartSpace } from "@office-open/core";

// Charts are created via @office-open/core and embedded in both docx and pptx
const chart = ChartSpace.create({
    type: "bar",
    data: [{ category: "A", value: 10 }, { category: "B", value: 20 }],
});
```

## PPTX Features

### Shapes with Text

```ts
new Shape({
    x: 50, y: 100, width: 300, height: 80,
    text: "Hello",
    fill: "4472C4",
    outline: { color: "0D47A1", width: 2 },
});
```

### Shapes with Effects

```ts
new Shape({
    x: 50, y: 100, width: 200, height: 120,
    text: "Shadow",
    fill: "ED7D31",
    effects: {
        outerShadow: { blur: 50800, distance: 38100, direction: 5400000, color: "000000", alpha: 50 },
        glow: { radius: 152400, color: "92D050", alpha: 60 },
        reflection: { blurRadius: 6350, distance: 38100, startAlpha: 90, endAlpha: 0 },
        softEdge: { radius: 50800 },
        rotation3D: { x: 20, y: 30, z: 10, perspective: 500 },
        extrusionH: 50000,
        material: "plastic",
        bevelTop: { width: 8, height: 8 },
    },
});
```

### Fill Types

```ts
// Solid
fill: "4472C4"
fill: { type: "solid", color: "4472C4" }

// Gradient
fill: { type: "gradient", stops: [{ position: 0, color: "1565C0" }, { position: 100, color: "42A5F5" }], angle: 45 }

// Image
fill: { type: "blip", data: imageBuffer, imageType: "png" }

// None
fill: { type: "noFill" }
```

## Measurement Units

| Unit       | Used For                    | Conversion                |
| ---------- | --------------------------- | ------------------------- |
| Twips      | Margins, spacing            | 1 inch = 1440 twips      |
| EMU        | Images, shapes, offsets     | 1 inch = 914400 EMU      |
| Half-points| Font sizes                  | Size 24 = 12pt            |
| 60000ths°  | Angles, rotation, effects   | 45° = 2700000             |

## References

See the `references/` directory for detailed API documentation on each topic.
