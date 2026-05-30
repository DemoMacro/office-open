# @office-open/pptx

> Generate and parse .pptx files with JS/TS with a declarative API.

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
# npm
npm install @office-open/pptx

# pnpm
pnpm add @office-open/pptx
```

## Quick Start

```typescript
import { Presentation, Shape, Packer, Paragraph, Run } from "@office-open/pptx";
import { writeFileSync } from "node:fs";

const pres = new Presentation({
  slides: [
    {
      children: [
        new Shape({
          text: "Hello World",
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

Read existing `.pptx` files and re-create them as `ISlideOptions[]`:

```typescript
import { parsePresentation, Presentation, Packer } from "@office-open/pptx";
import { readFileSync, writeFileSync } from "node:fs";

const slides = parsePresentation(new Uint8Array(readFileSync("input.pptx")));

// Modify parsed data, then re-generate
const pres = new Presentation({ slides });
const buffer = await Packer.toBuffer(pres);
writeFileSync("output.pptx", buffer);
```

## Benchmark

Performance vs [PptxGenJS](https://github.com/gitbrent/PptxGenJS) (higher ops/s is better, Windows 11 / Node 24).

DEFLATE = compressed (default), STORE = no compression. PptxGenJS always uses STORE.

**Create + toBuffer (end-to-end)**

| Scenario           | DEFLATE sync |  STORE sync | DEFLATE async | STORE async |   PptxGenJS |
| ------------------ | -----------: | ----------: | ------------: | ----------: | ----------: |
| Simple (2 shapes)  |    405 ops/s | 6,826 ops/s |     442 ops/s | 5,447 ops/s | 1,173 ops/s |
| Styled shapes (20) |    442 ops/s | 3,145 ops/s |     651 ops/s | 3,190 ops/s |   944 ops/s |
| Table (10×5)       |    642 ops/s | 4,719 ops/s |     830 ops/s | 5,072 ops/s |   765 ops/s |
| Full featured      |    448 ops/s | 2,305 ops/s |     621 ops/s | 2,543 ops/s |   601 ops/s |

**Large Files — Create + toBuffer**

| Scenario              | DEFLATE sync |  STORE sync | DEFLATE async | STORE async |   PptxGenJS |
| --------------------- | -----------: | ----------: | ------------: | ----------: | ----------: |
| 30 slides × 20 shapes |   65.9 ops/s | 123.7 ops/s |    75.0 ops/s | 147.2 ops/s | 111.4 ops/s |
| 100×10 table          |  207.3 ops/s | 418.6 ops/s |    27.9 ops/s | 500.0 ops/s | 107.6 ops/s |
| 50 slides full        |   62.6 ops/s | 157.7 ops/s |    66.6 ops/s | 163.0 ops/s |  94.3 ops/s |

**Large File (~100MB) — Mixed Content**

100 slides × (5 shapes + 2 unique images + 5×5 table). Speedup is vs PptxGenJS.

| Method        |      Speed |  Speedup |
| ------------- | ---------: | -------: |
| DEFLATE sync  | 1.63 ops/s |     4.1x |
| STORE sync    | 1.64 ops/s |     4.1x |
| DEFLATE async | 3.21 ops/s |     8.0x |
| STORE async   | 3.59 ops/s | **9.0x** |
| PptxGenJS     | 0.40 ops/s |          |

## Examples

Check the [demo folder](./demo) for working examples covering every feature.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
