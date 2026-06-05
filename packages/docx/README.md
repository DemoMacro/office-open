# @office-open/docx

![npm version](https://img.shields.io/npm/v/@office-open/docx)
![npm downloads](https://img.shields.io/npm/dw/@office-open/docx)
![npm license](https://img.shields.io/npm/l/@office-open/docx)

> Generate, parse, and patch .docx documents with a declarative TypeScript API. Works in Node.js and browsers.

## Features

- 📄 **Document Generation** — Create Word documents with sections, headers, footers, and page numbers
- ✍️ **Paragraphs & Text** — Rich text support with bold, italic, underline, strikethrough, and more
- 📊 **Tables** — Full table support with merged cells, borders, and styles
- 🖼️ **Images** — Inline and floating images with sizing, positioning, and wrapping
- 🔗 **Hyperlinks** — External and internal hyperlinks with custom styling
- 📑 **Headers & Footers** — First, last, even/odd page headers and footers
- 📋 **Lists** — Numbered and bulleted lists with multiple levels and custom formats
- 🎨 **Styles** — Paragraph, character, and table styles with inheritance
- 📖 **Table of Contents** — Auto-generated table of contents with custom styling
- 📝 **Footnotes & Endnotes** — Comprehensive footnote and endnote support
- 📈 **Charts** — Bar, line, pie, area, and scatter charts with customization
- 🔢 **Math Equations** — Full mathematical equation support via Office MathML (OMML)
- 🧩 **SmartArt** — Built-in SmartArt graphic generation
- 📚 **Bibliography** — Source management and citation support
- 💬 **Comments** — Document comments with author and date tracking
- 📝 **Track Revisions** — Insertions, deletions, and formatting changes
- 📋 **Content Controls** — Structured document tags (SDT) for form-like documents
- 📦 **Text Boxes** — Floating text boxes with content and styling
- ☑️ **Checkboxes** — Form checkbox support in documents
- 🖌️ **DrawingML** — Shapes with fills, shadows, effects, and transformations
- 🔤 **Custom Fonts** — Font embedding and custom font tables
- 🔧 **Template Patching** — Patch existing DOCX templates via placeholder replacement
- ⚙️ **Settings** — Comprehensive document settings and compatibility options

## Installation

```bash
# pnpm
pnpm add @office-open/docx

# npm
npm install @office-open/docx

# yarn
yarn add @office-open/docx

# bun
bun add @office-open/docx
```

## Quick Start

```typescript
import { Document, Paragraph, TextRun, Packer } from "@office-open/docx";
import { writeFileSync } from "node:fs";

const doc = new Document({
  sections: [
    {
      children: [
        new Paragraph({
          children: [
            new TextRun("Hello World"),
            new TextRun({
              text: " - Bold text",
              bold: true,
            }),
          ],
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
writeFileSync("My Document.docx", buffer);
```

## Examples

Check the [demo folder](./demo) for 100+ working examples covering every feature.

## Benchmark

Performance comparison against original `docx` (9.6.1) package (higher ops/s is better, Windows 11 / Node 24).

**Default** = XML DEFLATE level 1 (SuperFast, matching MS Office) + media STORE. **All STORE** = no compression (`{ compression: { xml: 0 } }`). **docx** (async) always uses DEFLATE for ALL entries including images (via JSZip, hardcoded, no STORE option).

```typescript
// Default (matches MS Office)
await Packer.toBuffer(doc);
// All STORE (no compression)
await Packer.toBuffer(doc, { compression: { xml: 0 } });
```

**Create + toBuffer (end-to-end)**

| Scenario                       | Default sync | Default async | All STORE sync | All STORE async |      docx |
| ------------------------------ | -----------: | ------------: | -------------: | --------------: | --------: |
| Simple (2p + 1 img)            |    174 ops/s |     155 ops/s |      202 ops/s |       203 ops/s |  82 ops/s |
| Styled paragraphs (20) + 1 img |    179 ops/s |     146 ops/s |      198 ops/s |       198 ops/s |  86 ops/s |
| Table (10x5)                   |    949 ops/s |     565 ops/s |    1,830 ops/s |     1,867 ops/s | 190 ops/s |
| Full featured + 2 imgs         |     94 ops/s |      89 ops/s |      102 ops/s |       100 ops/s |  57 ops/s |

**Large Files — Create + toBuffer**

| Scenario                       | Default sync | Default async | All STORE sync | All STORE async |       docx |
| ------------------------------ | -----------: | ------------: | -------------: | --------------: | ---------: |
| 2000 paragraphs + 20 images    |   4.68 ops/s |    3.13 ops/s |     3.50 ops/s |      3.08 ops/s | 2.07 ops/s |
| 200x10 table                   |   64.9 ops/s |    62.6 ops/s |     69.7 ops/s |      72.6 ops/s | 19.5 ops/s |
| 20 sections x 100p + 40 images |   1.58 ops/s |    1.69 ops/s |     1.79 ops/s |      1.65 ops/s | 1.16 ops/s |

**Large File (~100MB) — Mixed Content**

500 styled paragraphs + 38 mixed-size images (1-5MB, 100MB total) + 50x10 table. Speedup is vs docx.

| Method    |      Speed |  Speedup |
| --------- | ---------: | -------: |
| All STORE | 0.42 ops/s | **1.4x** |
| Default   | 0.31 ops/s |     1.1x |
| docx      | 0.29 ops/s |          |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
