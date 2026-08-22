# @office-open/docx

![npm version](https://img.shields.io/npm/v/@office-open/docx)
![npm downloads](https://img.shields.io/npm/dw/@office-open/docx)
![npm license](https://img.shields.io/npm/l/@office-open/docx)

> Create Word documents (.docx) in TypeScript and JavaScript — generate, parse, and patch from plain JSON, no Microsoft Office required. Built for AI agents and hand-written code alike; runs in Node.js, browsers, Deno, and Bun.

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

Performance vs the original `docx` (9.6.1) package (higher ops/s is better, Windows 11; each scenario runs under both Node 24 and Bun 1.4).

**Default** = XML DEFLATE level 1; media matches MS Office Word — already-compressed formats (PNG/JPEG/GIF) are STOREd, the rest (EMF/WMF/BMP/TIFF/…) use DEFLATE level 6. **All STORE** = no compression (`{ compression: { xml: 0, media: 0 } }`). The `docx` package (async only) always DEFLATEs every entry including images, with no STORE option.

```typescript
// Default (matches MS Office)
await generateDocument(options);
// All STORE (no compression)
await generateDocument(options, { compression: { xml: 0, media: 0 } });
// Stream as ReadableStream<Uint8Array> (pipe to a file / HTTP response)
generateDocumentStream(options);
```

**Create + toBuffer / toStream**

| Scenario                       | Runtime | Default sync | Default async | All STORE sync | All STORE async | Default stream | docx       |
| ------------------------------ | ------- | ------------ | ------------- | -------------- | --------------- | -------------- | ---------- |
| Simple (2p + 1 img)            | Node 24 | 905 ops/s    | 1,311 ops/s   | 2,478 ops/s    | 2,985 ops/s     | 10.1 ops/s     | 91.2 ops/s |
|                                | Bun 1.4 | 2,471 ops/s  | 758 ops/s     | 5,323 ops/s    | 5,246 ops/s     | 17.4 ops/s     | 63.7 ops/s |
| Styled paragraphs (20) + 1 img | Node 24 | 1,065 ops/s  | 1,405 ops/s   | 2,765 ops/s    | 2,752 ops/s     | 14.2 ops/s     | 91.1 ops/s |
|                                | Bun 1.4 | 2,994 ops/s  | 692 ops/s     | 4,567 ops/s    | 4,627 ops/s     | 17.4 ops/s     | 72.2 ops/s |
| Table (10x5)                   | Node 24 | 1,301 ops/s  | 1,286 ops/s   | 3,154 ops/s    | 3,163 ops/s     | 14.3 ops/s     | 196 ops/s  |
|                                | Bun 1.4 | 2,791 ops/s  | 691 ops/s     | 5,107 ops/s    | 4,882 ops/s     | 17.3 ops/s     | 252 ops/s  |
| Full featured + 2 imgs         | Node 24 | 763 ops/s    | 1,031 ops/s   | 1,671 ops/s    | 1,576 ops/s     | 12.7 ops/s     | 54.2 ops/s |
|                                | Bun 1.4 | 1,646 ops/s  | 575 ops/s     | 2,630 ops/s    | 2,456 ops/s     | 15.0 ops/s     | 41.1 ops/s |

**Large Files — Create + toBuffer / toStream**

| Scenario                       | Runtime | Default sync | Default async | All STORE sync | All STORE async | Default stream | docx       |
| ------------------------------ | ------- | ------------ | ------------- | -------------- | --------------- | -------------- | ---------- |
| 2000 paragraphs + 20 images    | Node 24 | 104.5 ops/s  | 109.4 ops/s   | 111.6 ops/s    | 114.9 ops/s     | 11.3 ops/s     | 2.85 ops/s |
|                                | Bun 1.4 | 148.8 ops/s  | 126.5 ops/s   | 168.0 ops/s    | 164.9 ops/s     | 14.4 ops/s     | 2.18 ops/s |
| 200x10 table                   | Node 24 | 252.4 ops/s  | 247.3 ops/s   | 282.1 ops/s    | 288.7 ops/s     | 13.7 ops/s     | 35.6 ops/s |
|                                | Bun 1.4 | 372.0 ops/s  | 272.4 ops/s   | 441.5 ops/s    | 405.6 ops/s     | 15.0 ops/s     | 44.5 ops/s |
| 20 sections x 100p + 40 images | Node 24 | 89.6 ops/s   | 94.5 ops/s    | 103.6 ops/s    | 103.7 ops/s     | 3.53 ops/s     | 1.73 ops/s |
|                                | Bun 1.4 | 141.9 ops/s  | 90.9 ops/s    | 163.7 ops/s    | 170.4 ops/s     | 3.40 ops/s     | 1.15 ops/s |

**Large File (~100MB) — Mixed Content**

500 styled paragraphs + 38 mixed-size images (1-5MB, 100MB total) + 50x10 table.

| Scenario                 | Runtime | Default sync | Default async | All STORE sync | All STORE async | Default stream | docx       |
| ------------------------ | ------- | ------------ | ------------- | -------------- | --------------- | -------------- | ---------- |
| Mixed (500p+38img+50x10) | Node 24 | 23.1 ops/s   | 21.2 ops/s    | 23.7 ops/s     | 23.9 ops/s      | 3.56 ops/s     | 0.28 ops/s |
|                          | Bun 1.4 | 38.8 ops/s   | 32.5 ops/s    | 36.5 ops/s     | 30.7 ops/s      | 4.09 ops/s     | 0.21 ops/s |

**Stream** = `generateDocumentStream` (default compression, fully drained) — it targets piping and a flat memory profile, not peak throughput.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
