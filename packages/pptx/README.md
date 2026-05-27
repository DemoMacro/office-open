# @office-open/pptx

> Generate and parse .pptx files with JS/TS with a declarative API.

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
- **Parsing** — Parse existing .pptx files with `parsePresentation` for round-trip workflows

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

## Performance

Performance vs [PptxGenJS](https://github.com/gitbrent/PptxGenJS) (higher hz is better, Windows 11 / Node 22):

**Object Creation (no pack)**

| Scenario                  | @office-open/pptx |  PptxGenJS |  Speedup |
| ------------------------- | ----------------: | ---------: | -------: |
| Simple (2 shapes)         |       4.69M ops/s | 573K ops/s | **8.2x** |
| Styled shapes (20 shapes) |        290K ops/s |  82K ops/s | **3.5x** |
| Table (10x5)              |       2.24M ops/s |  28K ops/s |  **80x** |
| Full featured             |        195K ops/s |  18K ops/s |  **11x** |

**Create + toBuffer (end-to-end)**

@office-open/pptx defaults to DEFLATE compression (smaller files). PptxGenJS always uses STORE (no compression). Speedup is based on same-compression (STORE vs STORE) comparison.

| Scenario                  | @office-open/pptx (DEFLATE) | @office-open/pptx (STORE) |   PptxGenJS |  Speedup |
| ------------------------- | --------------------------: | ------------------------: | ----------: | -------: |
| Simple (2 shapes)         |                   388 ops/s |               1,951 ops/s | 1,372 ops/s | **1.4x** |
| Styled shapes (20 shapes) |                   417 ops/s |               1,644 ops/s | 1,164 ops/s | **1.4x** |
| Table (10x5)              |                   534 ops/s |               1,905 ops/s |   918 ops/s | **2.1x** |
| Full featured             |                   489 ops/s |               1,114 ops/s |   710 ops/s | **1.6x** |

**Large Files — Create + toBuffer**

| Scenario              | @office-open/pptx (DEFLATE) | @office-open/pptx (STORE) | PptxGenJS |  Speedup |
| --------------------- | --------------------------: | ------------------------: | --------: | -------: |
| 30 slides × 20 shapes |                    72 ops/s |                 118 ops/s | 122 ops/s |     1.0x |
| 100×10 table          |                   165 ops/s |                 239 ops/s | 125 ops/s | **1.9x** |
| 50 slides full        |                    60 ops/s |                 135 ops/s | 106 ops/s | **1.3x** |

## Examples

Check the [demo folder](./demo) for working examples covering every feature.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
