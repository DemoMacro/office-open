# @office-open/pptx

![npm version](https://img.shields.io/npm/v/@office-open/pptx)
![npm downloads](https://img.shields.io/npm/dw/@office-open/pptx)
![npm license](https://img.shields.io/npm/l/@office-open/pptx)

> Create PowerPoint presentations (.pptx) in TypeScript and JavaScript — generate, parse, and patch from plain JSON, no Microsoft Office required. Built for AI agents and hand-written code alike; runs in Node.js, browsers, Deno, and Bun.

## Features

- 🖥️ **Slide Management** — Create presentations with multiple slides, slide masters, and slide layouts
- 🔷 **Shapes** — Rectangles, ellipses, lines, connectors, and custom geometry shapes
- ✍️ **Text & Rich Formatting** — Paragraphs, runs, fonts, colors, alignment, and spacing
- 📊 **Tables** — Full table support with rows, cells, borders, and cell properties
- 📈 **Charts** — Bar, line, pie, area, and scatter charts with customization
- 🧩 **SmartArt** — Built-in SmartArt graphic generation with multiple layouts and styles
- 🖼️ **Images** — Inline pictures with fills, transformations, and effects
- 🎨 **Backgrounds** — Solid color, gradient, and picture backgrounds
- 🔄 **Transitions** — Slide transitions with various types and durations
- ✨ **Animations** — Entrance, emphasis, exit, and motion path animations
- 🎬 **Media** — Video and audio embedding
- 🔗 **Hyperlinks** — Clickable hyperlinks on shapes and text
- 📑 **Headers & Footers** — Slide header/footer with date, slide number
- 📝 **Notes** — Speaker notes for each slide
- 👥 **Group Shapes** — Group multiple shapes together
- 🖌️ **DrawingML** — Shapes with fills, outlines, shadows, glow, reflection, and 3D effects
- 📖 **Parsing** — Parse existing .pptx files with `parsePresentation` for round-trip workflows
- 🔧 **Template Patching** — Patch existing PPTX templates via placeholder replacement

## Installation

```bash
# pnpm
pnpm add @office-open/pptx

# npm
npm install @office-open/pptx

# yarn
yarn add @office-open/pptx

# bun
bun add @office-open/pptx
```

## Quick Start

```typescript
import { generatePresentation } from "@office-open/pptx";
import { writeFileSync } from "node:fs";

const buffer = await generatePresentation({
  slides: [
    {
      children: [
        {
          shape: {
            textBody: {
              children: [{ paragraph: { children: ["Hello World"] } }],
            },
            fill: "4472C4",
            x: 100,
            y: 100,
            width: 600,
            height: 400,
          },
        },
      ],
    },
  ],
});

writeFileSync("presentation.pptx", buffer);
```

## Parsing

Read existing `.pptx` files and re-create them as `PresentationOptions`:

```typescript
import { parsePresentation, generatePresentation } from "@office-open/pptx";
import { readFileSync, writeFileSync } from "node:fs";

const opts = parsePresentation(new Uint8Array(readFileSync("input.pptx")));

// Modify parsed data, then re-generate
const buffer = await generatePresentation(opts);
writeFileSync("output.pptx", buffer);
```

## Benchmark

Performance vs [PptxGenJS](https://github.com/gitbrent/PptxGenJS) (higher ops/s is better, Windows 11 / Node 24).

**Default** = XML DEFLATE level 1 (SuperFast); media is split by type, matching MS Office PowerPoint — already-compressed formats (PNG/JPEG/GIF) are STOREd, the rest (EMF/WMF/BMP/TIFF/…) use DEFLATE level 6 / Normal (verified on a real MS Office file). **All STORE** = no compression (`{ compression: { xml: 0, media: 0 } }`). **PptxGenJS** (async only) defaults to STORE (via JSZip), supports DEFLATE via `compression: true` (applies to ALL entries including images).

```typescript
// Default (matches MS Office)
await generatePresentation(options);
// All STORE (no compression)
await generatePresentation(options, { compression: { xml: 0, media: 0 } });
// Stream as ReadableStream<Uint8Array> (pipe to a file / HTTP response)
generatePresentationStream(options);
```

**Create + toBuffer / toStream (end-to-end)**

| Scenario           | Default sync | Default async | All STORE sync | All STORE async | Default stream | PptxGenJS DEFLATE | PptxGenJS STORE |
| ------------------ | -----------: | ------------: | -------------: | --------------: | -------------: | ----------------: | --------------: |
| Simple (2 shapes)  |    944 ops/s |   1,155 ops/s |    2,761 ops/s |     3,068 ops/s |     13.2 ops/s |         184 ops/s |       190 ops/s |
| Styled shapes (20) |  1,044 ops/s |   1,210 ops/s |    2,861 ops/s |     2,788 ops/s |     13.5 ops/s |         175 ops/s |       181 ops/s |
| Table (10x5)       |  1,150 ops/s |   1,122 ops/s |    2,950 ops/s |     2,933 ops/s |     13.2 ops/s |         985 ops/s |       996 ops/s |
| Full featured      |    914 ops/s |     963 ops/s |    1,790 ops/s |     1,738 ops/s |     13.2 ops/s |        97.4 ops/s |      98.8 ops/s |

**Large Files — Create + toBuffer / toStream**

| Scenario              | Default sync | Default async | All STORE sync | All STORE async | Default stream | PptxGenJS DEFLATE | PptxGenJS STORE |
| --------------------- | -----------: | ------------: | -------------: | --------------: | -------------: | ----------------: | --------------: |
| 30 slides x 20 shapes |    198 ops/s |     205 ops/s |      369 ops/s |       370 ops/s |     2.91 ops/s |         125 ops/s |       131 ops/s |
| 30 slides x 10 images |    115 ops/s |     117 ops/s |      153 ops/s |       149 ops/s |     2.83 ops/s |        0.31 ops/s |      0.32 ops/s |
| 100x10 table          |    274 ops/s |     302 ops/s |      328 ops/s |       349 ops/s |     12.6 ops/s |         136 ops/s |       127 ops/s |
| 50 slides full        |   77.7 ops/s |    79.7 ops/s |      106 ops/s |       101 ops/s |     1.75 ops/s |        0.94 ops/s |      0.93 ops/s |

**Large File (~100MB) — Mixed Content**

40 slides x (2 shapes + 2 mixed-size images + 3x3 table).

| Scenario        | Default sync | Default async | All STORE sync | All STORE async | Default stream | PptxGenJS DEFLATE | PptxGenJS STORE |
| --------------- | -----------: | ------------: | -------------: | --------------: | -------------: | ----------------: | --------------: |
| 40 slides mixed |   24.2 ops/s |    21.2 ops/s |     24.8 ops/s |      25.4 ops/s |     1.68 ops/s |        0.23 ops/s |      0.23 ops/s |

**Stream column** = `generatePresentationStream` (default compression, fully drained). Streaming trades throughput for pipeability: each part is compressed in a Web Worker as it is produced, so a fixed per-part worker handoff dominates small decks — streaming targets memory footprint and piping (file / HTTP response), not peak ops/s.

## Examples

Check the [demo folder](./demo) for working examples covering every feature.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
