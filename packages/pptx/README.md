# @office-open/pptx

> Generate and parse .pptx presentations with a declarative TypeScript API. Works in Node.js and browsers.

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
| Simple (2 shapes)  |    919 ops/s |     549 ops/s |    2,731 ops/s |     2,767 ops/s |     12.8 ops/s |         172 ops/s |       173 ops/s |
| Styled shapes (20) |    935 ops/s |     558 ops/s |    2,567 ops/s |     2,345 ops/s |     12.6 ops/s |         174 ops/s |       176 ops/s |
| Table (10x5)       |  1,093 ops/s |     602 ops/s |    2,882 ops/s |     2,976 ops/s |     12.8 ops/s |         919 ops/s |       940 ops/s |
| Full featured      |    727 ops/s |     533 ops/s |    1,703 ops/s |     1,682 ops/s |     12.5 ops/s |        93.3 ops/s |      92.1 ops/s |

**Large Files — Create + toBuffer / toStream**

| Scenario              | Default sync | Default async | All STORE sync | All STORE async | Default stream | PptxGenJS DEFLATE | PptxGenJS STORE |
| --------------------- | -----------: | ------------: | -------------: | --------------: | -------------: | ----------------: | --------------: |
| 30 slides x 20 shapes |    160 ops/s |     109 ops/s |      267 ops/s |       259 ops/s |     2.81 ops/s |         121 ops/s |       121 ops/s |
| 30 slides x 10 images |    105 ops/s |    78.9 ops/s |      135 ops/s |       138 ops/s |     2.75 ops/s |        0.30 ops/s |      0.30 ops/s |
| 100x10 table          |    259 ops/s |     213 ops/s |      273 ops/s |       292 ops/s |     10.7 ops/s |         123 ops/s |       134 ops/s |
| 50 slides full        |   67.7 ops/s |    49.7 ops/s |     89.8 ops/s |      86.8 ops/s |     1.78 ops/s |        0.92 ops/s |      0.91 ops/s |

**Large File (~100MB) — Mixed Content**

40 slides x (2 shapes + 2 mixed-size images + 3x3 table).

| Scenario        | Default sync | Default async | All STORE sync | All STORE async | Default stream | PptxGenJS DEFLATE | PptxGenJS STORE |
| --------------- | -----------: | ------------: | -------------: | --------------: | -------------: | ----------------: | --------------: |
| 40 slides mixed |   21.7 ops/s |    20.2 ops/s |     23.2 ops/s |      23.9 ops/s |     1.67 ops/s |        0.22 ops/s |      0.21 ops/s |

**Stream column** = `generatePresentationStream` (default compression, fully drained). Streaming trades throughput for pipeability: each part is compressed in a Web Worker as it is produced, so a fixed per-part worker handoff dominates small decks — streaming targets memory footprint and piping (file / HTTP response), not peak ops/s.

## Examples

Check the [demo folder](./demo) for working examples covering every feature.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
