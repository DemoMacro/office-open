# @office-open/pptx

> Generate .pptx files with JS/TS with a declarative API.

## Installation

```bash
# npm
npm install @office-open/pptx

# pnpm
pnpm add @office-open/pptx
```

## Quick Start

```typescript
import { Presentation, Shape, Packer } from "@office-open/pptx";
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

## Features

- **Shapes** — Rectangles, ellipses, lines, connectors, group shapes
- **Text** — Paragraphs, runs, fonts, colors, alignment, bullets
- **Tables** — Rows, cells, borders, cell spans, vertical align
- **Charts** — Bar, line, pie, area, scatter
- **SmartArt** — Multiple layouts and styles
- **Media** — Video and audio embedding
- **DrawingML** — Fills (solid, gradient, pattern), outlines, effects (shadow, glow, reflection, 3D)
- **Backgrounds** — Solid color, gradient
- **Transitions & Animations** — Slide transitions and shape animations
- **Hyperlinks** — On shapes and text runs
- **Headers & Footers** — Date, slide number
- **Notes** — Speaker notes per slide

## Examples

See the [demo folder](https://github.com/DemoMacro/office-open/tree/main/packages/pptx/demo) for working examples.

## License

[MIT](https://github.com/DemoMacro/office-open/blob/main/LICENSE) &copy; [Demo Macro](https://imst.xyz/)
