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
| Simple (2 shapes)  |    400 ops/s | 5,057 ops/s |     526 ops/s | 5,691 ops/s | 1,192 ops/s |
| Styled shapes (20) |    548 ops/s | 3,497 ops/s |     703 ops/s | 3,319 ops/s |   989 ops/s |
| Table (10×5)       |    692 ops/s | 4,199 ops/s |     696 ops/s | 4,075 ops/s |   840 ops/s |
| Full featured      |    597 ops/s | 2,584 ops/s |     620 ops/s | 2,546 ops/s |   683 ops/s |

**Large Files — Create + toBuffer**

| Scenario              | DEFLATE sync |  STORE sync | DEFLATE async | STORE async |   PptxGenJS |
| --------------------- | -----------: | ----------: | ------------: | ----------: | ----------: |
| 30 slides × 20 shapes |   70.0 ops/s | 172.6 ops/s |    91.5 ops/s | 214.5 ops/s | 101.2 ops/s |
| 100×10 table          |  201.3 ops/s | 491.4 ops/s |    31.1 ops/s | 555.2 ops/s | 117.4 ops/s |
| 50 slides full        |   76.1 ops/s | 220.5 ops/s |    71.2 ops/s | 203.9 ops/s |  79.8 ops/s |

**Large File (~100MB) — Mixed Content**

100 slides × (5 shapes + 2 unique images + 5×5 table). Speedup is vs PptxGenJS.

| Method        |      Speed |  Speedup |
| ------------- | ---------: | -------: |
| DEFLATE sync  | 1.79 ops/s |     3.4x |
| STORE sync    | 1.87 ops/s |     3.6x |
| DEFLATE async | 3.87 ops/s |     7.4x |
| STORE async   | 4.09 ops/s | **7.9x** |
| PptxGenJS     | 0.52 ops/s |          |

## Examples

Check the [demo folder](./demo) for working examples covering every feature.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
