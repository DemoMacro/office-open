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

| Scenario           | Default sync | Default async | All STORE sync | All STORE async | PptxGenJS |
| ------------------ | -----------: | ------------: | -------------: | --------------: | --------: |
| Simple (2 shapes)  |  1,154 ops/s |     561 ops/s |    2,788 ops/s |     2,236 ops/s | 176 ops/s |
| Styled shapes (20) |    944 ops/s |     563 ops/s |    2,274 ops/s |     2,379 ops/s | 143 ops/s |
| Table (10x5)       |  1,747 ops/s |     659 ops/s |    5,305 ops/s |     5,196 ops/s | 659 ops/s |
| Full featured      |    762 ops/s |     482 ops/s |    1,605 ops/s |     1,451 ops/s |  75 ops/s |

**Large Files — Create + toBuffer**

| Scenario              | Default sync | Default async | All STORE sync | All STORE async |  PptxGenJS |
| --------------------- | -----------: | ------------: | -------------: | --------------: | ---------: |
| 30 slides x 20 shapes |    133 ops/s |      85 ops/s |      211 ops/s |       242 ops/s |   91 ops/s |
| 30 slides x 10 images |   10.1 ops/s |    9.88 ops/s |     10.2 ops/s |      10.1 ops/s | 0.35 ops/s |
| 100x10 table          |    502 ops/s |     369 ops/s |      698 ops/s |       704 ops/s |  119 ops/s |
| 50 slides full        |   25.1 ops/s |    20.3 ops/s |     26.9 ops/s |      26.1 ops/s | 1.02 ops/s |

**Large File (~100MB) — Mixed Content**

40 slides x (2 shapes + 2 mixed-size images + 3x3 table). Speedup is vs PptxGenJS.

| Method    |      Speed | Speedup |
| --------- | ---------: | ------: |
| Default   |  8.0 ops/s |   26.7x |
| All STORE |  8.2 ops/s |   27.4x |
| PptxGenJS | 0.30 ops/s |         |

## Examples

Check the [demo folder](./demo) for working examples covering every feature.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
