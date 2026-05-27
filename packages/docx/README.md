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

Performance comparison against original `docx` (9.6.1) package (Windows 11 / Node 22):

**Object Creation (no pack)**

| Scenario                                                | @office-open/docx |       docx |  Speedup |
| ------------------------------------------------------- | ----------------: | ---------: | -------: |
| Simple document (2 paragraphs)                          |       20.0K ops/s | 6.2K ops/s | **3.2x** |
| Styled paragraphs (20 paragraphs)                       |       11.4K ops/s | 4.6K ops/s | **2.5x** |
| Table (10x5 cells)                                      |       10.1K ops/s | 3.7K ops/s | **2.7x** |
| Full featured (header/footer/headings/table/paragraphs) |        5.5K ops/s | 2.8K ops/s | **2.0x** |

**Create + toBuffer (end-to-end)**

Both libraries use DEFLATE compression.

| Scenario                                                | @office-open/docx |      docx |  Speedup |
| ------------------------------------------------------- | ----------------: | --------: | -------: |
| Simple document (2 paragraphs)                          |         388 ops/s | 194 ops/s | **2.0x** |
| Styled paragraphs (20 paragraphs)                       |         402 ops/s | 232 ops/s | **1.7x** |
| Table (10x5 cells)                                      |         463 ops/s | 238 ops/s | **1.9x** |
| Full featured (header/footer/headings/table/paragraphs) |         346 ops/s | 179 ops/s | **1.9x** |

**Large Files — Create + toBuffer**

| Scenario                     | @office-open/docx |       docx |  Speedup |
| ---------------------------- | ----------------: | ---------: | -------: |
| 2000 paragraphs              |        42.2 ops/s | 27.7 ops/s | **1.5x** |
| 200×10 table                 |        77.4 ops/s | 37.3 ops/s | **2.1x** |
| 20 sections × 100 paragraphs |        47.0 ops/s | 30.9 ops/s | **1.5x** |

**Large File (~100MB) — Mixed Content**

2000 paragraphs + 200 unique images (500KB each) + 100×10 table. Speedup is vs docx.

| Method                    |      Speed |  Speedup |
| ------------------------- | ---------: | -------: |
| @office-open/docx (sync)  | 0.36 ops/s |     1.2x |
| @office-open/docx (async) | 0.40 ops/s | **1.3x** |
| docx                      | 0.30 ops/s |          |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
