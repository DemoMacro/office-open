# PPTX API Reference

Complete API reference for `@office-open/pptx`.

## Presentation Structure

```
Presentation
├── title, creator, subject, description
├── slides: ISlideOptions[]
│   ├── children: (Shape | Table | Image | GroupShape | Connector)[]
│   └── background?: IFillOptions
```

## Shapes

### Basic Shape

```ts
import { Shape } from "@office-open/pptx";

new Shape({
    id: 1,                      // Optional, auto-assigned
    name: "My Shape",           // Optional display name
    x: 50,                      // Position X (pixels)
    y: 100,                     // Position Y (pixels)
    width: 300,                 // Width (pixels)
    height: 80,                 // Height (pixels)
    text: "Hello World",        // Simple text content
    fill: "4472C4",             // Solid fill shorthand
});
```

### Shape with Paragraphs

```ts
import { Shape, Paragraph, TextRun } from "@office-open/pptx";

new Shape({
    x: 50, y: 100, width: 400, height: 120,
    paragraphs: [
        new Paragraph({
            children: [
                new TextRun({ text: "Bold title", bold: true, fontSize: 24, color: "FFFFFF" }),
            ],
            alignment: "center",
        }),
        new Paragraph({
            children: [
                new TextRun({ text: "Subtitle", fontSize: 16, color: "CCCCCC" }),
            ],
        }),
    ],
    fill: { type: "solid", color: "2E74B5" },
});
```

### Shape Text Options

```ts
new Shape({
    text: "Content",
    textAnchor: "ctr",          // "ctr" | "t" | "b"
    textVertical: false,        // Vertical text direction
    textAutoFit: "normal",      // "none" | "normal" | "shape"
    textWrap: "square",         // "square" | "none"
    textMargins: {              // EMU values
        top: 45720, bottom: 45720,
        left: 91440, right: 91440,
    },
    textColumns: 2,             // Number of text columns
    textColumnSpacing: 45720,   // Spacing between columns (EMU)
});
```

### Geometry Presets

```ts
new Shape({
    geometry: "rect",           // Default
    // "roundRect", "ellipse", "triangle", "diamond", "hexagon",
    // "star5", "star6", "arrow", "chevron", "cloud", "heart",
    // "lightningBolt", "moon", "sun", and many more
});
```

## Fills

### Solid Fill

```ts
fill: "4472C4"                                   // Shorthand
fill: { type: "solid", color: "2E74B5" }         // Explicit
```

### Gradient Fill

```ts
fill: {
    type: "gradient",
    stops: [
        { position: 0, color: "0D47A1" },
        { position: 50, color: "1976D2" },
        { position: 100, color: "42A5F5" },
    ],
    angle: 45,                // Degrees
    // path: "circle",        // For radial gradient
}
```

### Image Fill

```ts
import { readFileSync } from "node:fs";

fill: {
    type: "blip",
    data: readFileSync("texture.png"),
    imageType: "png",
}
```

### No Fill

```ts
fill: { type: "noFill" }
```

## Outline

```ts
outline: {
    color: "1565C0",
    width: 2,                    // Points
    dashStyle: "solid",          // "solid" | "dash" | "dashDot" | "dot" | "lgDash"
    capStyle: "flat",            // "flat" | "round" | "square"
    compound: "single",          // "single" | "double" | "thickThin"
}
```

## Effects

### Outer Shadow

```ts
effects: {
    outerShadow: {
        blur: 50800,             // Blur radius (EMU)
        distance: 38100,         // Offset distance (EMU)
        direction: 5400000,      // Angle (60000ths of a degree)
        color: "000000",
        alpha: 50,               // Opacity (0-100)
        rotateWithShape: true,
    },
}
```

### Inner Shadow

```ts
effects: {
    innerShadow: {
        blur: 40000,
        distance: 30000,
        direction: 5400000,
        color: "000000",
        alpha: 40,
    },
}
```

### Glow

```ts
effects: {
    glow: {
        radius: 152400,          // Glow radius (EMU)
        color: "92D050",
        alpha: 60,               // Opacity (0-100)
    },
}
```

### Reflection

```ts
effects: {
    reflection: {
        blurRadius: 6350,
        distance: 38100,
        direction: 5400000,
        startAlpha: 90,          // Start opacity (0-100)
        endAlpha: 0,             // End opacity (0-100)
        startPosition: 0,        // Start position (0-100)
        endPosition: 100,        // End position (0-100)
    },
}
```

### Soft Edge

```ts
effects: {
    softEdge: { radius: 50800 },
}
```

### 3D Rotation

```ts
effects: {
    rotation3D: {
        x: 20,                   // X rotation (degrees)
        y: 30,                   // Y rotation (degrees)
        z: 10,                   // Z rotation (degrees)
        perspective: 500,        // Perspective value
    },
}
```

### Extrusion & Bevel

```ts
effects: {
    rotation3D: { x: 25, y: 15 },
    extrusionH: 50000,           // Extrusion height (EMU)
    material: "plastic",         // "plastic" | "metal" | "matte" | "warmMatte" | "flat" | "powder"
    bevelTop: { width: 8, height: 8 },    // Bevel dimensions (points)
    bevelBottom: { width: 5, height: 5 },
    lighting: "threePt",         // "flat" | "threePt" | "balanced" | "soft" | "harsh" | ...
}
```

## Tables

```ts
import { Table, TableRow, TableCell } from "@office-open/pptx";

new Table({
    x: 50, y: 200, width: 600, height: 200,
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    text: "Header 1",
                    fill: "4472C4",
                    bold: true,
                    color: "FFFFFF",
                }),
                new TableCell({ text: "Header 2" }),
            ],
        }),
    ],
});
```

## Images

```ts
import { Image } from "@office-open/pptx";

new Image({
    x: 100, y: 100, width: 300, height: 200,
    data: imageBuffer,
    imageType: "png",
});
```

## Charts

Charts are created via `@office-open/core` and embedded in slides.

## Connectors & Lines

```ts
import { Connector } from "@office-open/pptx";

new Connector({
    x: 100, y: 100, width: 200, height: 0,
    outline: { color: "000000", width: 2 },
});
```

## Group Shapes

```ts
import { GroupShape } from "@office-open/pptx";

new GroupShape({
    x: 50, y: 50, width: 500, height: 300,
    children: [
        new Shape({ x: 0, y: 0, width: 240, height: 140, text: "A", fill: "4472C4" }),
        new Shape({ x: 260, y: 0, width: 240, height: 140, text: "B", fill: "ED7D31" }),
    ],
});
```

## Slide Layout & Background

```ts
new Presentation({
    slides: [{
        background: {
            type: "solid",
            color: "FFFFFF",
        },
        children: [...],
    }],
});
```

## Transitions

```ts
new Presentation({
    slides: [{
        transition: {
            type: "fade",
            duration: 1000,
        },
        children: [...],
    }],
});
```

## Animations

```ts
new Shape({
    x: 50, y: 100, width: 200, height: 80,
    text: "Animated",
    animation: {
        type: "appear",
        trigger: "onClick",
        delay: 0,
        duration: 500,
    },
});
```

## Parsing

```ts
import { Packer, Presentation } from "@office-open/pptx";
import { readFileSync } from "node:fs";

const buffer = readFileSync("input.pptx");
const pres = await Packer.fromBuffer(buffer);
const parsed = pres.toParsedDocument();
```
