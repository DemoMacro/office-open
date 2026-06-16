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
import { generateDocument } from "@office-open/docx";
import { writeFileSync } from "node:fs";

const buffer = await generateDocument({
  sections: [
    {
      children: [
        {
          paragraph: {
            children: ["Hello World", { text: " - Bold text", bold: true }],
          },
        },
      ],
    },
  ],
});

writeFileSync("My Document.docx", buffer);
```

## Examples

Check the [demo folder](./demo) for 100+ working examples covering every feature.

## Benchmark

Performance vs original `docx` (9.6.1) package (higher ops/s is better, Windows 11 / Node 24).

**Default** = XML DEFLATE level 1 (SuperFast); media is split by type, matching MS Office Word — already-compressed formats (PNG/JPEG/GIF) are STOREd, the rest (EMF/WMF/BMP/TIFF/…) use DEFLATE level 1 (verified on a real MS Office file). **All STORE** = no compression (`{ compression: { xml: 0, media: 0 } }`). **docx** (async only) always uses DEFLATE for ALL entries including images (via JSZip, hardcoded, no STORE option).

```typescript
// Default (matches MS Office)
await generateDocument(options);
// All STORE (no compression)
await generateDocument(options, { compression: { xml: 0, media: 0 } });
```

**Create + toBuffer (end-to-end)**

| Scenario                       | Default sync | Default async | All STORE sync | All STORE async |      docx |
| ------------------------------ | -----------: | ------------: | -------------: | --------------: | --------: |
| Simple (2p + 1 img)            |  1,047 ops/s |     605 ops/s |    1,444 ops/s |     1,448 ops/s |  82 ops/s |
| Styled paragraphs (20) + 1 img |  1,133 ops/s |     593 ops/s |    1,610 ops/s |     1,679 ops/s |  89 ops/s |
| Table (10x5)                   |  1,695 ops/s |     785 ops/s |    4,144 ops/s |     3,928 ops/s | 239 ops/s |
| Full featured + 2 imgs         |    769 ops/s |     439 ops/s |      785 ops/s |       814 ops/s |  57 ops/s |

**Large Files — Create + toBuffer**

| Scenario                       | Default sync | Default async | All STORE sync | All STORE async |       docx |
| ------------------------------ | -----------: | ------------: | -------------: | --------------: | ---------: |
| 2000 paragraphs + 20 images    |   62.7 ops/s |    51.9 ops/s |     34.7 ops/s |      35.2 ops/s |  3.0 ops/s |
| 200x10 table                   |    261 ops/s |     218 ops/s |      271 ops/s |       271 ops/s | 36.3 ops/s |
| 20 sections x 100p + 40 images |   35.8 ops/s |    27.1 ops/s |     18.0 ops/s |      18.7 ops/s |  1.8 ops/s |

**Large File (~100MB) — Mixed Content**

500 styled paragraphs + 38 mixed-size images (1-5MB, 100MB total) + 50x10 table.

| Scenario                 | Default sync | Default async | All STORE sync | All STORE async |       docx |
| ------------------------ | -----------: | ------------: | -------------: | --------------: | ---------: |
| Mixed (500p+38img+50x10) |    9.7 ops/s |     8.2 ops/s |      5.1 ops/s |       4.2 ops/s | 0.30 ops/s |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
