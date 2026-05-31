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
| Simple (2 paragraphs)                                   |    349 ops/s | 2,085 ops/s |     379 ops/s | 2,495 ops/s | 192 ops/s |
| Styled paragraphs (20)                                  |    477 ops/s | 1,983 ops/s |     511 ops/s | 1,918 ops/s | 236 ops/s |
| Table (10×5)                                            |    448 ops/s | 1,505 ops/s |     480 ops/s | 1,565 ops/s | 203 ops/s |
| Full featured (header/footer/headings/table/paragraphs) |    326 ops/s |   992 ops/s |     380 ops/s | 1,100 ops/s | 163 ops/s |

**Large Files — Create + toBuffer**

| Scenario                     | DEFLATE sync |  STORE sync | DEFLATE async | STORE async |       docx |
| ---------------------------- | -----------: | ----------: | ------------: | ----------: | ---------: |
| 2000 paragraphs              |   43.6 ops/s |  53.4 ops/s |    17.7 ops/s |  53.8 ops/s | 20.0 ops/s |
| 200×10 table                 |   78.0 ops/s | 106.6 ops/s |    24.1 ops/s | 109.6 ops/s | 30.0 ops/s |
| 20 sections × 100 paragraphs |   44.0 ops/s |  70.9 ops/s |    21.6 ops/s |  75.9 ops/s | 22.9 ops/s |

**Large File (~100MB) — Mixed Content**

2000 paragraphs + 200 unique images (500KB each) + 100×10 table. Speedup is vs docx.

| Method        |      Speed |  Speedup |
| ------------- | ---------: | -------: |
| DEFLATE sync  | 0.37 ops/s | **1.8x** |
| STORE sync    | 0.37 ops/s | **1.7x** |
| DEFLATE async | 0.41 ops/s | **2.0x** |
| STORE async   | 0.42 ops/s | **2.0x** |
| docx          | 0.21 ops/s |          |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
