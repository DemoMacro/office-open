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
| Simple (2 shapes)  |    943 ops/s |     527 ops/s |    2,579 ops/s |     2,633 ops/s |     13.5 ops/s |         183 ops/s |       195 ops/s |
| Styled shapes (20) |    931 ops/s |     540 ops/s |    2,698 ops/s |     2,465 ops/s |     13.4 ops/s |         184 ops/s |       186 ops/s |
| Table (10x5)       |  1,127 ops/s |     627 ops/s |    3,109 ops/s |     3,087 ops/s |     13.5 ops/s |         970 ops/s |     1,033 ops/s |
| Full featured      |    890 ops/s |     499 ops/s |    1,614 ops/s |     1,709 ops/s |     13.4 ops/s |        96.8 ops/s |       102 ops/s |

**Large Files — Create + toBuffer / toStream**

| Scenario              | Default sync | Default async | All STORE sync | All STORE async | Default stream | PptxGenJS DEFLATE | PptxGenJS STORE |
| --------------------- | -----------: | ------------: | -------------: | --------------: | -------------: | ----------------: | --------------: |
| 30 slides x 20 shapes |    177 ops/s |     111 ops/s |      279 ops/s |       279 ops/s |     2.94 ops/s |         117 ops/s |       130 ops/s |
| 30 slides x 10 images |    110 ops/s |    73.5 ops/s |      143 ops/s |       139 ops/s |     2.82 ops/s |        0.34 ops/s |      0.34 ops/s |
| 100x10 table          |    273 ops/s |     225 ops/s |      303 ops/s |       317 ops/s |     13.1 ops/s |         123 ops/s |       137 ops/s |
| 50 slides full        |   70.8 ops/s |    51.3 ops/s |     95.4 ops/s |      90.0 ops/s |     1.85 ops/s |        1.00 ops/s |      1.01 ops/s |

**Large File (~100MB) — Mixed Content**

40 slides x (2 shapes + 2 mixed-size images + 3x3 table).

| Scenario        | Default sync | Default async | All STORE sync | All STORE async | Default stream | PptxGenJS DEFLATE | PptxGenJS STORE |
| --------------- | -----------: | ------------: | -------------: | --------------: | -------------: | ----------------: | --------------: |
| 40 slides mixed |   23.3 ops/s |    21.1 ops/s |     24.8 ops/s |      24.5 ops/s |     1.68 ops/s |        0.24 ops/s |      0.24 ops/s |

**Stream column** = `generatePresentationStream` (default compression, fully drained). Streaming trades throughput for pipeability: each part is compressed in a Web Worker as it is produced, so a fixed per-part worker handoff dominates small decks — streaming targets memory footprint and piping (file / HTTP response), not peak ops/s.

## Examples

Check the [demo folder](./demo) for working examples covering every feature.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
