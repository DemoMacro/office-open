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
import { generatePresentation } from "@office-open/pptx";
import { writeFileSync } from "node:fs";

const buffer = await generatePresentation({
  slides: [
    {
      children: [
        {
          shape: {
            textBody: {
              children: [{ paragraph: { children: ["Hello World"] } }],
            },
            fill: "4472C4",
            x: 100,
            y: 100,
            width: 600,
            height: 400,
          },
        },
      ],
    },
  ],
});

writeFileSync("presentation.pptx", buffer);
```

## Parsing

Read existing `.pptx` files and re-create them as `PresentationOptions`:

```typescript
import { parsePresentation, generatePresentation } from "@office-open/pptx";
import { readFileSync, writeFileSync } from "node:fs";

const opts = parsePresentation(new Uint8Array(readFileSync("input.pptx")));

// Modify parsed data, then re-generate
const buffer = await generatePresentation(opts);
writeFileSync("output.pptx", buffer);
```

## Benchmark

Performance vs [PptxGenJS](https://github.com/gitbrent/PptxGenJS) (higher ops/s is better, Windows 11 / Node 24).

**Default** = XML DEFLATE level 1 (SuperFast, matching MS Office) + media STORE. **All STORE** = no compression (`{ compression: { xml: 0 } }`). **PptxGenJS** (async only) defaults to STORE (via JSZip), supports DEFLATE via `compression: true` (applies to ALL entries including images).

```typescript
// Default (matches MS Office)
await generatePresentation(options);
// All STORE (no compression)
await generatePresentation(options, { compression: { xml: 0 } });
```

**Create + toBuffer (end-to-end)**

| Scenario           | Default sync | Default async | All STORE sync | All STORE async | PptxGenJS DEFLATE | PptxGenJS STORE |
| ------------------ | -----------: | ------------: | -------------: | --------------: | ----------------: | --------------: |
| Simple (2 shapes)  |  1,754 ops/s |     748 ops/s |    6,289 ops/s |     6,449 ops/s |         194 ops/s |       196 ops/s |
| Styled shapes (20) |  1,450 ops/s |     731 ops/s |    4,735 ops/s |     5,520 ops/s |         186 ops/s |       182 ops/s |
| Table (10x5)       |  1,633 ops/s |     773 ops/s |    7,426 ops/s |     6,907 ops/s |         863 ops/s |       909 ops/s |
| Full featured      |  1,546 ops/s |     767 ops/s |    4,752 ops/s |     4,927 ops/s |         102 ops/s |       104 ops/s |

**Large Files — Create + toBuffer**

| Scenario              | Default sync | Default async | All STORE sync | All STORE async | PptxGenJS DEFLATE | PptxGenJS STORE |
| --------------------- | -----------: | ------------: | -------------: | --------------: | ----------------: | --------------: |
| 30 slides x 20 shapes |    259 ops/s |     138 ops/s |      501 ops/s |       508 ops/s |         118 ops/s |       121 ops/s |
| 30 slides x 10 images |    333 ops/s |     165 ops/s |      840 ops/s |       826 ops/s |        0.34 ops/s |      0.34 ops/s |
| 100x10 table          |    623 ops/s |     455 ops/s |      980 ops/s |       982 ops/s |         118 ops/s |       124 ops/s |
| 50 slides full        |    165 ops/s |      95 ops/s |      334 ops/s |       318 ops/s |        1.00 ops/s |      0.95 ops/s |

**Large File (~100MB) — Mixed Content**

40 slides x (2 shapes + 2 mixed-size images + 3x3 table).

| Scenario        | Default sync | Default async | All STORE sync | All STORE async | PptxGenJS DEFLATE | PptxGenJS STORE |
| --------------- | -----------: | ------------: | -------------: | --------------: | ----------------: | --------------: |
| 40 slides mixed |    221 ops/s |     120 ops/s |      641 ops/s |       688 ops/s |        0.24 ops/s |      0.24 ops/s |

## Examples

Check the [demo folder](./demo) for working examples covering every feature.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
