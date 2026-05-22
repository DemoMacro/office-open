# @office-open/pptx

> Generate .pptx files with JS/TS with a declarative API.

## Features

- **Slide Management** — Create presentations with multiple slides, slide masters, and slide layouts
- **Shapes** — Rectangles, ellipses, lines, connectors, and custom geometry shapes
- **Text & Rich Formatting** — Paragraphs, runs, fonts, colors, alignment, and spacing
- **Tables** — Full table support with rows, cells, borders, and cell properties
- **Charts** — Bar, line, pie, area, and scatter charts with customization
- **SmartArt** — Built-in SmartArt graphic generation with multiple layouts and styles
- **Images** — Inline pictures with fills, transformations, and effects
- **Backgrounds** — Solid color, gradient, and picture backgrounds
- **Transitions** — Slide transitions with various types and durations
- **Animations** — Entrance, emphasis, exit, and motion path animations
- **Media** — Video and audio embedding
- **Hyperlinks** — Clickable hyperlinks on shapes and text
- **Headers & Footers** — Slide header/footer with date, slide number
- **Notes** — Speaker notes for each slide
- **Group Shapes** — Group multiple shapes together
- **DrawingML** — Shapes with fills, outlines, shadows, glow, reflection, and 3D effects

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

## Performance

Performance vs [PptxGenJS](https://github.com/gitbrent/PptxGenJS) (higher hz is better, Windows 11 / Node 22):

**Object Creation (no pack)**

| Scenario                  | @office-open/pptx |   PptxGenJS |   Speedup |
| ------------------------- | ----------------: | ----------: | --------: |
| Simple (2 shapes)         |       5.06M ops/s |  659K ops/s |  **7.7x** |
| Styled shapes (20 shapes) |        244K ops/s | 82.9K ops/s |  **2.9x** |
| Table (10x5)              |       2.32M ops/s | 28.9K ops/s | **80.2x** |
| Full featured             |        199K ops/s | 18.1K ops/s | **11.0x** |

**Create + toBuffer (end-to-end)**

| Scenario                  | @office-open/pptx | PptxGenJS (DEFLATE) |  Speedup |
| ------------------------- | ----------------: | ------------------: | -------: |
| Simple (2 shapes)         |       1,705 ops/s |         1,456 ops/s | **1.2x** |
| Styled shapes (20 shapes) |       1,377 ops/s |         1,243 ops/s | **1.1x** |
| Table (10x5)              |       1,644 ops/s |           983 ops/s | **1.7x** |
| Full featured             |       1,138 ops/s |           806 ops/s | **1.4x** |

**Large Files — Create + toBuffer**

| Scenario              | @office-open/pptx | PptxGenJS (DEFLATE) |  Speedup |
| --------------------- | ----------------: | ------------------: | -------: |
| 10 slides × 10 shapes |         493 ops/s |           475 ops/s | **1.0x** |
| 50×10 table           |         450 ops/s |           243 ops/s | **1.9x** |
| 20 slides full        |         450 ops/s |           343 ops/s | **1.3x** |

## Examples

Check the [demo folder](./demo) for working examples covering every feature.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
