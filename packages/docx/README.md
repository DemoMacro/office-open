# @office-open/docx

![npm version](https://img.shields.io/npm/v/@office-open/docx)
![npm downloads](https://img.shields.io/npm/dw/@office-open/docx)
![npm license](https://img.shields.io/npm/l/@office-open/docx)

> Easily generate .docx files with JS/TS with a nice declarative API. Works for Node and on the Browser.

## Features

- **Document Generation** - Create Word documents with sections, headers, footers, and page numbers
- **Paragraphs & Text** - Rich text support with bold, italic, underline, strikethrough, and more
- **Tables** - Full table support with merged cells, borders, and styles
- **Images** - Inline and floating images with sizing, positioning, and wrapping
- **Hyperlinks** - External and internal hyperlinks with custom styling
- **Headers & Footers** - First, last, even/odd page headers and footers
- **Lists** - Numbered and bulleted lists with multiple levels and custom formats
- **Styles** - Paragraph, character, and table styles with inheritance
- **Table of Contents** - Auto-generated table of contents with custom styling
- **Footnotes & Endnotes** - Comprehensive footnote and endnote support
- **Charts** - Bar, line, pie, area, and scatter charts with customization
- **Math Equations** - Full mathematical equation support via MathML
- **SmartArt** - Built-in SmartArt graphic generation
- **Bibliography** - Source management and citation support
- **Comments** - Document comments with author and date tracking
- **Track Revisions** - Insertions, deletions, and formatting changes
- **Content Controls** - Structured document tags (SDT) for form-like documents
- **Text Boxes** - Floating text boxes with content and styling
- **Checkboxes** - Form checkbox support in documents
- **DrawingML** - Shapes with fills, shadows, effects, and transformations
- **Custom Fonts** - Font embedding and custom font tables
- **Template Patching** - Patch existing DOCX templates via placeholder replacement
- **Settings** - Comprehensive document settings and compatibility options

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

Performance comparison against original `docx` (9.6.1) package (Windows 11 / Node 24).

Both libraries use DEFLATE compression.

**Create + toBuffer (end-to-end)**

| Scenario                                                | @office-open/docx |      docx |  Speedup |
| ------------------------------------------------------- | ----------------: | --------: | -------: |
| Simple document (2 paragraphs)                          |         341 ops/s | 173 ops/s | **2.0x** |
| Styled paragraphs (20 paragraphs)                       |         341 ops/s | 210 ops/s | **1.6x** |
| Table (10x5 cells)                                      |         419 ops/s | 216 ops/s | **1.9x** |
| Full featured (header/footer/headings/table/paragraphs) |         369 ops/s | 168 ops/s | **2.2x** |

**Large Files — Create + toBuffer**

| Scenario                     | @office-open/docx |       docx |  Speedup |
| ---------------------------- | ----------------: | ---------: | -------: |
| 2000 paragraphs              |        39.3 ops/s | 24.2 ops/s | **1.6x** |
| 200×10 table                 |        78.1 ops/s | 30.5 ops/s | **2.6x** |
| 20 sections × 100 paragraphs |        42.5 ops/s | 25.6 ops/s | **1.7x** |

**Large File (~100MB) — Mixed Content**

2000 paragraphs + 200 unique images (500KB each) + 100×10 table. Speedup is vs docx.

| Method                    |      Speed |  Speedup |
| ------------------------- | ---------: | -------: |
| @office-open/docx (sync)  | 0.35 ops/s |     1.2x |
| @office-open/docx (async) | 0.38 ops/s | **1.3x** |
| docx                      | 0.29 ops/s |          |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
