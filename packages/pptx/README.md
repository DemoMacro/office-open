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
- **Template Patching** — Patch existing PPTX templates via placeholder replacement

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

Performance vs [PptxGenJS](https://github.com/gitbrent/PptxGenJS) (higher hz is better, Windows 11 / Node 24).

@office-open/pptx defaults to DEFLATE compression (smaller files). PptxGenJS always uses STORE (no compression). Speedup is based on same-compression (STORE vs STORE) comparison.

**Create + toBuffer (end-to-end)**

| Scenario                  | @office-open/pptx (DEFLATE) | @office-open/pptx (STORE) |   PptxGenJS |  Speedup |
| ------------------------- | --------------------------: | ------------------------: | ----------: | -------: |
| Simple (2 shapes)         |                   505 ops/s |               6,213 ops/s | 1,033 ops/s | **6.0x** |
| Styled shapes (20 shapes) |                   451 ops/s |               3,217 ops/s |   971 ops/s | **3.3x** |
| Table (10x5)              |                   606 ops/s |               3,715 ops/s |   759 ops/s | **4.9x** |
| Full featured             |                   513 ops/s |               1,882 ops/s |   607 ops/s | **3.1x** |

**Large Files — Create + toBuffer**

| Scenario              | @office-open/pptx (DEFLATE) | @office-open/pptx (STORE) |  PptxGenJS |  Speedup |
| --------------------- | --------------------------: | ------------------------: | ---------: | -------: |
| 30 slides × 20 shapes |                  69.7 ops/s |                 129 ops/s |  109 ops/s | **1.2x** |
| 100×10 table          |                   156 ops/s |                 235 ops/s |  116 ops/s | **2.0x** |
| 50 slides full        |                  56.9 ops/s |                 143 ops/s | 95.8 ops/s | **1.5x** |

**Large File (~100MB) — Mixed Content**

100 slides × (5 shapes + 2 unique images + 5×5 table). Speedup is vs PptxGenJS.

| Method                            |      Speed |  Speedup |
| --------------------------------- | ---------: | -------: |
| @office-open/pptx (DEFLATE sync)  | 1.76 ops/s |     3.9x |
| @office-open/pptx (DEFLATE async) | 3.29 ops/s | **7.3x** |
| @office-open/pptx (STORE sync)    | 1.81 ops/s |     4.0x |
| PptxGenJS                         | 0.45 ops/s |          |

## Examples

Check the [demo folder](./demo) for working examples covering every feature.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
