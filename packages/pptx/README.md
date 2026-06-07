# @office-open/pptx

> Generate and parse .pptx presentations with a declarative TypeScript API. Works in Node.js and browsers.

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
# pnpm
pnpm add @office-open/pptx

# npm
npm install @office-open/pptx

# yarn
yarn add @office-open/pptx

# bun
bun add @office-open/pptx
```

## Quick Start

```typescript
import { Presentation, Shape, Packer, Paragraph, TextRun } from "@office-open/pptx";
import { writeFileSync } from "node:fs";

const pres = new Presentation({
  slides: [
    {
      children: [
        new Shape({
          textBody: { children: [new Paragraph({ children: [new TextRun("Hello World")] })] },
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

Read existing `.pptx` files and re-create them as `PresentationOptions`:

```typescript
import { parsePresentation, Presentation, Packer } from "@office-open/pptx";
import { readFileSync, writeFileSync } from "node:fs";

const opts = parsePresentation(new Uint8Array(readFileSync("input.pptx")));

// Modify parsed data, then re-generate
const pres = new Presentation(opts);
const buffer = await Packer.toBuffer(pres);
writeFileSync("output.pptx", buffer);
```

## Benchmark

Performance vs [PptxGenJS](https://github.com/gitbrent/PptxGenJS) (higher ops/s is better, Windows 11 / Node 24).

**Default** = XML DEFLATE level 1 (SuperFast, matching MS Office) + media STORE. **All STORE** = no compression (`{ compression: { xml: 0 } }`). **PptxGenJS** (async only) defaults to STORE (via JSZip), supports DEFLATE via `compression: true` (applies to ALL entries including images).

```typescript
// Default (matches MS Office)
await Packer.toBuffer(pres);
// All STORE (no compression)
await Packer.toBuffer(pres, { compression: { xml: 0 } });
```

**Create + toBuffer (end-to-end)**

| Scenario           | Default sync | Default async | All STORE sync | All STORE async | PptxGenJS DEFLATE | PptxGenJS STORE |
| ------------------ | -----------: | ------------: | -------------: | --------------: | ----------------: | --------------: |
| Simple (2 shapes)  |  1,507 ops/s |     740 ops/s |    5,449 ops/s |     4,913 ops/s |         164 ops/s |       168 ops/s |
| Styled shapes (20) |  1,390 ops/s |     698 ops/s |    4,526 ops/s |     4,314 ops/s |         168 ops/s |       173 ops/s |
| Table (10x5)       |  1,538 ops/s |     735 ops/s |    5,905 ops/s |     6,394 ops/s |         775 ops/s |       770 ops/s |
| Full featured      |  1,325 ops/s |     532 ops/s |    3,369 ops/s |     3,995 ops/s |          92 ops/s |        87 ops/s |

**Large Files — Create + toBuffer**

| Scenario              | Default sync | Default async | All STORE sync | All STORE async | PptxGenJS DEFLATE | PptxGenJS STORE |
| --------------------- | -----------: | ------------: | -------------: | --------------: | ----------------: | --------------: |
| 30 slides x 20 shapes |    236 ops/s |     116 ops/s |      460 ops/s |       377 ops/s |         113 ops/s |       124 ops/s |
| 30 slides x 10 images |    309 ops/s |     164 ops/s |      818 ops/s |       777 ops/s |        0.29 ops/s |      0.28 ops/s |
| 100x10 table          |    555 ops/s |     441 ops/s |      799 ops/s |       876 ops/s |         108 ops/s |       115 ops/s |
| 50 slides full        |    150 ops/s |      93 ops/s |      317 ops/s |       310 ops/s |        0.84 ops/s |      0.83 ops/s |

**Large File (~100MB) — Mixed Content**

40 slides x (2 shapes + 2 mixed-size images + 3x3 table).

| Scenario        | Default sync | Default async | All STORE sync | All STORE async | PptxGenJS DEFLATE | PptxGenJS STORE |
| --------------- | -----------: | ------------: | -------------: | --------------: | ----------------: | --------------: |
| 40 slides mixed |    208 ops/s |     116 ops/s |      636 ops/s |       609 ops/s |        0.20 ops/s |      0.20 ops/s |

## Examples

Check the [demo folder](./demo) for working examples covering every feature.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
