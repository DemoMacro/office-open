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
import { Presentation, Shape, Packer, Paragraph, TextRun } from "@office-open/pptx";
import { writeFileSync } from "node:fs";

const pres = new Presentation({
  slides: [
    {
      children: [
        new Shape({
          textBody: { children: [new Paragraph({ children: [new TextRun("Hello World")] })] },
          fill: "4472C4",
          x: 100,
          y: 100,
          width: 600,
          height: 400,
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(pres);
writeFileSync("presentation.pptx", buffer);
```

## Parsing

Read existing `.pptx` files and re-create them as `PresentationOptions`:

```typescript
import { parsePresentation, Presentation, Packer } from "@office-open/pptx";
import { readFileSync, writeFileSync } from "node:fs";

const opts = parsePresentation(new Uint8Array(readFileSync("input.pptx")));

// Modify parsed data, then re-generate
const pres = new Presentation(opts);
const buffer = await Packer.toBuffer(pres);
writeFileSync("output.pptx", buffer);
```

## Benchmark

Performance vs [PptxGenJS](https://github.com/gitbrent/PptxGenJS) (higher ops/s is better, Windows 11 / Node 24).

**Default** = XML DEFLATE level 1 (SuperFast, matching MS Office) + media STORE. **All STORE** = no compression (`{ compression: { xml: 0 } }`). **PptxGenJS** (async only) defaults to STORE (via JSZip), supports DEFLATE via `compression: true` (applies to ALL entries including images).

```typescript
// Default (matches MS Office)
await Packer.toBuffer(pres);
// All STORE (no compression)
await Packer.toBuffer(pres, { compression: { xml: 0 } });
```

**Create + toBuffer (end-to-end)**

| Scenario           | Default sync | Default async | All STORE sync | All STORE async | PptxGenJS DEFLATE | PptxGenJS STORE |
| ------------------ | -----------: | ------------: | -------------: | --------------: | ----------------: | --------------: |
| Simple (2 shapes)  |  1,610 ops/s |     739 ops/s |    5,778 ops/s |     5,350 ops/s |         183 ops/s |       195 ops/s |
| Styled shapes (20) |  1,556 ops/s |     754 ops/s |    4,899 ops/s |     4,672 ops/s |         183 ops/s |       184 ops/s |
| Table (10x5)       |  1,561 ops/s |     758 ops/s |    6,590 ops/s |     6,517 ops/s |         837 ops/s |       912 ops/s |
| Full featured      |  1,484 ops/s |     756 ops/s |    4,650 ops/s |     4,705 ops/s |          98 ops/s |        99 ops/s |

**Large Files — Create + toBuffer**

| Scenario              | Default sync | Default async | All STORE sync | All STORE async | PptxGenJS DEFLATE | PptxGenJS STORE |
| --------------------- | -----------: | ------------: | -------------: | --------------: | ----------------: | --------------: |
| 30 slides x 20 shapes |    249 ops/s |     147 ops/s |      460 ops/s |       455 ops/s |         117 ops/s |       124 ops/s |
| 30 slides x 10 images |    309 ops/s |     161 ops/s |      796 ops/s |       784 ops/s |        0.32 ops/s |      0.32 ops/s |
| 100x10 table          |    623 ops/s |     450 ops/s |      907 ops/s |       890 ops/s |         118 ops/s |       120 ops/s |
| 50 slides full        |    151 ops/s |      90 ops/s |      300 ops/s |       293 ops/s |        0.93 ops/s |      0.93 ops/s |

**Large File (~100MB) — Mixed Content**

40 slides x (2 shapes + 2 mixed-size images + 3x3 table).

| Scenario        | Default sync | Default async | All STORE sync | All STORE async | PptxGenJS DEFLATE | PptxGenJS STORE |
| --------------- | -----------: | ------------: | -------------: | --------------: | ----------------: | --------------: |
| 40 slides mixed |    224 ops/s |     121 ops/s |      630 ops/s |       662 ops/s |        0.21 ops/s |      0.21 ops/s |

## Examples

Check the [demo folder](./demo) for working examples covering every feature.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
