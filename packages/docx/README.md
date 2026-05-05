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
| Simple document (2 paragraphs)                          |       27.6K ops/s | 6.9K ops/s | **4.0x** |
| Styled paragraphs (20 paragraphs)                       |       26.1K ops/s | 4.3K ops/s | **6.0x** |
| Table (10x5 cells)                                      |       18.3K ops/s | 3.6K ops/s | **5.1x** |
| Full featured (header/footer/headings/table/paragraphs) |       15.3K ops/s | 2.9K ops/s | **5.3x** |

**Create + toBuffer (end-to-end)**

| Scenario                                                | @office-open/docx |      docx |  Speedup |
| ------------------------------------------------------- | ----------------: | --------: | -------: |
| Simple document (2 paragraphs)                          |         398 ops/s | 252 ops/s | **1.6x** |
| Styled paragraphs (20 paragraphs)                       |         492 ops/s | 290 ops/s | **1.7x** |
| Table (10x5 cells)                                      |         442 ops/s | 268 ops/s | **1.6x** |
| Full featured (header/footer/headings/table/paragraphs) |         347 ops/s | 219 ops/s | **1.6x** |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
