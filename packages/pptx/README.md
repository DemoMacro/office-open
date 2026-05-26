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
| Simple (2 shapes)         |       4.60M ops/s | 545K ops/s | **8.4x** |
| Styled shapes (20 shapes) |        231K ops/s |  77K ops/s | **3.0x** |
| Table (10x5)              |       2.23M ops/s |  28K ops/s |  **79x** |
| Full featured             |        193K ops/s |  17K ops/s |  **11x** |

**Create + toBuffer (end-to-end)**

@office-open/pptx defaults to DEFLATE compression (smaller files). PptxGenJS always uses STORE (no compression). Speedup is based on same-compression (STORE vs STORE) comparison.

| Scenario                  | @office-open/pptx (DEFLATE) | @office-open/pptx (STORE) |   PptxGenJS |  Speedup |
| ------------------------- | --------------------------: | ------------------------: | ----------: | -------: |
| Simple (2 shapes)         |                   373 ops/s |               1,885 ops/s | 1,383 ops/s | **1.4x** |
| Styled shapes (20 shapes) |                   393 ops/s |               1,359 ops/s | 1,152 ops/s | **1.2x** |
| Table (10x5)              |                   532 ops/s |               1,498 ops/s |   898 ops/s | **1.7x** |
| Full featured             |                   455 ops/s |               1,057 ops/s |   752 ops/s | **1.4x** |

**Large Files — Create + toBuffer**

| Scenario              | @office-open/pptx (DEFLATE) | @office-open/pptx (STORE) | PptxGenJS |  Speedup |
| --------------------- | --------------------------: | ------------------------: | --------: | -------: |
| 10 slides × 10 shapes |                   207 ops/s |                 488 ops/s | 445 ops/s | **1.1x** |
| 50×10 table           |                   293 ops/s |                 458 ops/s | 228 ops/s | **2.0x** |
| 20 slides full        |                   167 ops/s |                 411 ops/s | 313 ops/s | **1.3x** |

## Examples

Check the [demo folder](./demo) for working examples covering every feature.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
