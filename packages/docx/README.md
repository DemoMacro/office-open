# @office-open/docx

![npm version](https://img.shields.io/npm/v/@office-open/docx)
![npm downloads](https://img.shields.io/npm/dw/@office-open/docx)
![npm license](https://img.shields.io/npm/l/@office-open/docx)

> Generate .docx files with a nice declarative API. Works for Node.js and on the Browser.

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

Performance comparison against original `docx` (9.6.1) package:

| Scenario                                                | @office-open/docx |     docx |   Speedup |
| ------------------------------------------------------- | ----------------: | -------: | --------: |
| Simple document (2 paragraphs)                          |          7,672 hz | 4,334 hz | **1.77x** |
| Styled paragraphs (20 paragraphs)                       |          6,301 hz | 4,157 hz | **1.52x** |
| Table (10x5 cells)                                      |          3,854 hz | 2,595 hz | **1.49x** |
| Full featured (header/footer/headings/table/paragraphs) |          2,745 hz | 2,366 hz | **1.16x** |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
