# @office-open/docx

![npm version](https://img.shields.io/npm/v/@office-open/docx)
![npm downloads](https://img.shields.io/npm/dw/@office-open/docx)
![npm license](https://img.shields.io/npm/l/@office-open/docx)

> Easily generate .docx files with JS/TS with a nice declarative API. Works for Node and on the Browser.

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
- 🔢 **Math Equations** — Full mathematical equation support via MathML
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
# Install with npm
$ npm install @office-open/docx

# Install with pnpm
$ pnpm add @office-open/docx
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

DEFLATE = compressed (default), STORE = no compression. Both libraries use DEFLATE.

**Create + toBuffer (end-to-end)**

| Scenario                                                | DEFLATE sync |  STORE sync | DEFLATE async | STORE async |      docx |
| ------------------------------------------------------- | -----------: | ----------: | ------------: | ----------: | --------: |
| Simple (2 paragraphs)                                   |    300 ops/s | 2,370 ops/s |     367 ops/s | 2,577 ops/s | 153 ops/s |
| Styled paragraphs (20)                                  |    449 ops/s | 2,145 ops/s |     520 ops/s | 2,047 ops/s | 208 ops/s |
| Table (10×5)                                            |    522 ops/s | 1,850 ops/s |     451 ops/s | 1,818 ops/s | 196 ops/s |
| Full featured (header/footer/headings/table/paragraphs) |    382 ops/s | 1,099 ops/s |     384 ops/s | 1,054 ops/s | 168 ops/s |

**Large Files — Create + toBuffer**

| Scenario                     | DEFLATE sync |  STORE sync | DEFLATE async | STORE async |       docx |
| ---------------------------- | -----------: | ----------: | ------------: | ----------: | ---------: |
| 2000 paragraphs              |   43.3 ops/s |  62.7 ops/s |    19.8 ops/s |  64.9 ops/s | 24.2 ops/s |
| 200×10 table                 |   85.5 ops/s | 107.0 ops/s |    23.0 ops/s | 110.7 ops/s | 34.3 ops/s |
| 20 sections × 100 paragraphs |   47.9 ops/s |  73.7 ops/s |    20.4 ops/s |  76.3 ops/s | 20.4 ops/s |

**Large File (~100MB) — Mixed Content**

2000 paragraphs + 200 unique images (500KB each) + 100×10 table. Speedup is vs docx.

| Method        |      Speed |  Speedup |
| ------------- | ---------: | -------: |
| DEFLATE sync  | 0.34 ops/s | **1.1x** |
| STORE sync    | 0.34 ops/s | **1.1x** |
| DEFLATE async | 0.38 ops/s | **1.2x** |
| STORE async   | 0.38 ops/s | **1.2x** |
| docx          | 0.31 ops/s |          |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
