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
| Simple (2p + 1 img)            |    868 ops/s |     562 ops/s |    2,334 ops/s |     2,312 ops/s |     13.9 ops/s | 71.4 ops/s |
| Styled paragraphs (20) + 1 img |    944 ops/s |     567 ops/s |    2,720 ops/s |     2,494 ops/s |     13.0 ops/s | 84.1 ops/s |
| Table (10x5)                   |  1,081 ops/s |     590 ops/s |    2,583 ops/s |     2,774 ops/s |     12.9 ops/s |  205 ops/s |
| Full featured + 2 imgs         |    777 ops/s |     488 ops/s |    1,433 ops/s |     1,645 ops/s |     11.1 ops/s | 49.7 ops/s |

**Large Files — Create + toBuffer / toStream**

| Scenario                       | Default sync | Default async | All STORE sync | All STORE async | Default stream |       docx |
| ------------------------------ | -----------: | ------------: | -------------: | --------------: | -------------: | ---------: |
| 2000 paragraphs + 20 images    |  107.9 ops/s |    83.9 ops/s |    112.9 ops/s |     105.3 ops/s |     8.52 ops/s | 2.66 ops/s |
| 200x10 table                   |  215.8 ops/s |   173.7 ops/s |    215.2 ops/s |     218.4 ops/s |     6.30 ops/s | 33.2 ops/s |
| 20 sections x 100p + 40 images |   83.8 ops/s |    67.4 ops/s |    103.4 ops/s |     103.8 ops/s |     3.14 ops/s | 1.67 ops/s |

**Large File (~100MB) — Mixed Content**

500 styled paragraphs + 38 mixed-size images (1-5MB, 100MB total) + 50x10 table.

| Scenario                 | Default sync | Default async | All STORE sync | All STORE async | Default stream |       docx |
| ------------------------ | -----------: | ------------: | -------------: | --------------: | -------------: | ---------: |
| Mixed (500p+38img+50x10) |   23.9 ops/s |    22.9 ops/s |     24.2 ops/s |      24.6 ops/s |     3.59 ops/s | 0.29 ops/s |

**Stream column** = `generateDocumentStream` (default compression, fully drained). Streaming trades throughput for pipeability: each part is compressed in a Web Worker as it is produced, so a fixed per-part worker handoff dominates small documents — streaming targets memory footprint and piping (file / HTTP response), not peak ops/s.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
