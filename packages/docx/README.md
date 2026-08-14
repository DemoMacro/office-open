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

**Default** = XML DEFLATE level 1 (SuperFast); media is split by type, matching MS Office Word — already-compressed formats (PNG/JPEG/GIF) are STOREd, the rest (EMF/WMF/BMP/TIFF/…) use DEFLATE level 6 / Normal (verified on a real MS Office file). **All STORE** = no compression (`{ compression: { xml: 0, media: 0 } }`). **docx** (async only) always uses DEFLATE for ALL entries including images (via JSZip, hardcoded, no STORE option).

```typescript
// Default (matches MS Office)
await generateDocument(options);
// All STORE (no compression)
await generateDocument(options, { compression: { xml: 0, media: 0 } });
```

**Create + toBuffer (end-to-end)**

| Scenario                       | Default sync | Default async | All STORE sync | All STORE async |       docx |
| ------------------------------ | -----------: | ------------: | -------------: | --------------: | ---------: |
| Simple (2p + 1 img)            |    938 ops/s |     599 ops/s |    2,279 ops/s |     2,299 ops/s | 89.1 ops/s |
| Styled paragraphs (20) + 1 img |    932 ops/s |     614 ops/s |    2,741 ops/s |     2,712 ops/s |  101 ops/s |
| Table (10x5)                   |  1,245 ops/s |     664 ops/s |    2,700 ops/s |     2,680 ops/s |  216 ops/s |
| Full featured + 2 imgs         |    794 ops/s |     508 ops/s |    1,579 ops/s |     1,627 ops/s | 57.6 ops/s |

**Large Files — Create + toBuffer**

| Scenario                       | Default sync | Default async | All STORE sync | All STORE async |       docx |
| ------------------------------ | -----------: | ------------: | -------------: | --------------: | ---------: |
| 2000 paragraphs + 20 images    |   85.2 ops/s |    75.4 ops/s |     90.2 ops/s |      86.7 ops/s | 2.98 ops/s |
| 200x10 table                   |    204 ops/s |     181 ops/s |      220 ops/s |       225 ops/s | 37.8 ops/s |
| 20 sections x 100p + 40 images |   81.2 ops/s |    68.7 ops/s |     89.9 ops/s |      89.7 ops/s | 1.81 ops/s |

**Large File (~100MB) — Mixed Content**

500 styled paragraphs + 38 mixed-size images (1-5MB, 100MB total) + 50x10 table.

| Scenario                 | Default sync | Default async | All STORE sync | All STORE async |       docx |
| ------------------------ | -----------: | ------------: | -------------: | --------------: | ---------: |
| Mixed (500p+38img+50x10) |   23.9 ops/s |    22.2 ops/s |     24.9 ops/s |      24.6 ops/s | 0.30 ops/s |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
