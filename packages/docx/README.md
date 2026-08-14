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
// Stream as ReadableStream<Uint8Array> (pipe to a file / HTTP response)
generateDocumentStream(options);
```

**Create + toBuffer / toStream (end-to-end)**

| Scenario                       | Default sync | Default async | All STORE sync | All STORE async | Default stream |       docx |
| ------------------------------ | -----------: | ------------: | -------------: | --------------: | -------------: | ---------: |
| Simple (2p + 1 img)            |    906 ops/s |     588 ops/s |    1,999 ops/s |     2,276 ops/s |     14.1 ops/s | 82.1 ops/s |
| Styled paragraphs (20) + 1 img |    987 ops/s |     651 ops/s |    2,571 ops/s |     2,502 ops/s |     14.4 ops/s | 94.8 ops/s |
| Table (10x5)                   |  1,193 ops/s |     675 ops/s |    2,704 ops/s |     2,627 ops/s |     14.2 ops/s |  188 ops/s |
| Full featured + 2 imgs         |    765 ops/s |     523 ops/s |    1,513 ops/s |     1,534 ops/s |     12.4 ops/s | 54.3 ops/s |

**Large Files — Create + toBuffer / toStream**

| Scenario                       | Default sync | Default async | All STORE sync | All STORE async | Default stream |       docx |
| ------------------------------ | -----------: | ------------: | -------------: | --------------: | -------------: | ---------: |
| 2000 paragraphs + 20 images    |   84.8 ops/s |    81.2 ops/s |     92.5 ops/s |      92.5 ops/s |     11.2 ops/s | 2.81 ops/s |
| 200x10 table                   |    191 ops/s |     173 ops/s |      207 ops/s |       230 ops/s |     13.8 ops/s | 34.9 ops/s |
| 20 sections x 100p + 40 images |   77.1 ops/s |    65.5 ops/s |     90.9 ops/s |      83.5 ops/s |     3.72 ops/s | 1.72 ops/s |

**Large File (~100MB) — Mixed Content**

500 styled paragraphs + 38 mixed-size images (1-5MB, 100MB total) + 50x10 table.

| Scenario                 | Default sync | Default async | All STORE sync | All STORE async | Default stream |       docx |
| ------------------------ | -----------: | ------------: | -------------: | --------------: | -------------: | ---------: |
| Mixed (500p+38img+50x10) |   23.5 ops/s |    22.0 ops/s |     23.7 ops/s |      23.1 ops/s |     3.60 ops/s | 0.29 ops/s |

**Stream column** = `generateDocumentStream` (default compression, fully drained). Streaming trades throughput for pipeability: each part is compressed in a Web Worker as it is produced, so a fixed per-part worker handoff dominates small documents — streaming targets memory footprint and piping (file / HTTP response), not peak ops/s.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
