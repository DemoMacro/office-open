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

Object instantiation performance vs [PptxGenJS](https://github.com/gitbrent/PptxGenJS) (higher hz is better, run on Windows 11 / Node 22):

| Scenario | @office-open/pptx | PptxGenJS | Speedup |
|---|---|---|---|
| Simple (2 shapes) | 4.64M ops/s | 559K ops/s | **8.3x** |
| Styled shapes (20 shapes) | 243K ops/s | 77K ops/s | **3.2x** |
| Table (10x5) | 1.86M ops/s | 27K ops/s | **69.8x** |
| Full featured | 188K ops/s | 19K ops/s | **9.7x** |

## Examples

Check the [demo folder](./demo) for 15 working examples covering every feature.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
