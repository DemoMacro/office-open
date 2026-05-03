# @office-open/pptx

![npm version](https://img.shields.io/npm/v/@office-open/pptx)
![npm downloads](https://img.shields.io/npm/dw/@office-open/pptx)
![npm license](https://img.shields.io/npm/l/@office-open/pptx)

> Easily generate .pptx files with JS/TS with a declarative API. Works for Node and on the Browser.

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
# Install with npm
$ npm install @office-open/pptx

# Install with pnpm
$ pnpm add @office-open/pptx
```

## Quick Start

```typescript
import { Presentation, Shape, SolidFill, Paragraph, Run, Packer } from "@office-open/pptx";
import { writeFileSync } from "node:fs";

const pres = new Presentation({
    slides: [
        {
            children: [
                new Shape({
                    textBody: new TextBody({
                        paragraphs: [
                            new Paragraph({
                                children: [new Run({ text: "Hello World", fontSize: 32 })],
                            }),
                        ],
                    }),
                    properties: {
                        fill: new SolidFill("4472C4"),
                        transform: { x: 100, y: 100, width: 600, height: 400 },
                    },
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(pres);
writeFileSync("presentation.pptx", buffer);
```

## Examples

Check the [demo folder](./demo) for 15 working examples covering every feature.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
