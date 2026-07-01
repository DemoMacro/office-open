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
```

**Create + toBuffer (end-to-end)**

| Scenario           | Default sync | Default async | All STORE sync | All STORE async | PptxGenJS DEFLATE | PptxGenJS STORE |
| ------------------ | -----------: | ------------: | -------------: | --------------: | ----------------: | --------------: |
| Simple (2 shapes)  |  1,176 ops/s |     646 ops/s |    4,345 ops/s |     3,969 ops/s |         185 ops/s |       199 ops/s |
| Styled shapes (20) |  1,182 ops/s |     660 ops/s |    4,197 ops/s |     4,224 ops/s |         196 ops/s |       194 ops/s |
| Table (10x5)       |  1,471 ops/s |     716 ops/s |    8,044 ops/s |     7,416 ops/s |         928 ops/s |     1,020 ops/s |
| Full featured      |  1,033 ops/s |     639 ops/s |    2,924 ops/s |     2,598 ops/s |         107 ops/s |       102 ops/s |

**Large Files — Create + toBuffer**

| Scenario              | Default sync | Default async | All STORE sync | All STORE async | PptxGenJS DEFLATE | PptxGenJS STORE |
| --------------------- | -----------: | ------------: | -------------: | --------------: | ----------------: | --------------: |
| 30 slides x 20 shapes |    255 ops/s |     143 ops/s |      567 ops/s |       547 ops/s |         127 ops/s |       135 ops/s |
| 30 slides x 10 images |    120 ops/s |    84.4 ops/s |      169 ops/s |       163 ops/s |        0.34 ops/s |      0.34 ops/s |
| 100x10 table          |    603 ops/s |     446 ops/s |    1,053 ops/s |     1,052 ops/s |         135 ops/s |       121 ops/s |
| 50 slides full        |   86.7 ops/s |    54.4 ops/s |      126 ops/s |       128 ops/s |        1.01 ops/s |      1.02 ops/s |

**Large File (~100MB) — Mixed Content**

40 slides x (2 shapes + 2 mixed-size images + 3x3 table).

| Scenario        | Default sync | Default async | All STORE sync | All STORE async | PptxGenJS DEFLATE | PptxGenJS STORE |
| --------------- | -----------: | ------------: | -------------: | --------------: | ----------------: | --------------: |
| 40 slides mixed |   24.5 ops/s |    22.4 ops/s |     25.3 ops/s |      25.0 ops/s |        0.24 ops/s |      0.24 ops/s |

## Examples

Check the [demo folder](./demo) for working examples covering every feature.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
