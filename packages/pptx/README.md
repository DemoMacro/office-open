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

| Scenario                  | @office-open/pptx | PptxGenJS   | Speedup   |
| ------------------------- | ----------------- | ----------- | --------- |
| Simple (2 shapes)         | 4.70M ops/s       | 1.20M ops/s | **3.9x**  |
| Styled shapes (20 shapes) | 256K ops/s        | 166K ops/s  | **1.5x**  |
| Table (10x5)              | 2.23M ops/s       | 35K ops/s   | **62.9x** |
| Full featured             | 198K ops/s        | 24K ops/s   | **8.4x**  |

**Create + toBuffer (end-to-end)**

| Scenario                  | @office-open/pptx | PptxGenJS (DEFLATE) | Speedup  |
| ------------------------- | ----------------- | ------------------- | -------- |
| Simple (2 shapes)         | 1,922 ops/s       | 1,152 ops/s         | **1.7x** |
| Styled shapes (20 shapes) | 1,540 ops/s       | 1,123 ops/s         | **1.4x** |
| Table (10x5)              | 1,790 ops/s       | 934 ops/s           | **1.9x** |
| Full featured             | 1,125 ops/s       | 827 ops/s           | **1.4x** |

## Examples

Check the [demo folder](./demo) for 15 working examples covering every feature.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
