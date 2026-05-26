# DOCX API Reference

Complete API reference for `@office-open/docx`.

## Document Structure

```
Document
├── sections: SectionOptions[]
│   ├── properties: SectionPropertiesOptionsBase
│   │   ├── page: { size, margins, columns, borders }
│   │   └── type: SectionType (CONTINUOUS | NEXT_PAGE | EVEN_PAGE | ODD_PAGE)
│   ├── headers: { default, first, even, odd }
│   ├── footers: { default, first, even, odd }
│   └── children: (Paragraph | Table | ImageRun | ...)[]
```

## Text Formatting

### TextRun Options

```ts
new TextRun({
  text: "Hello",
  bold: true,
  italics: true,
  underline: { type: "single", color: "FF0000" },
  strike: "single", // "single" | "double" | "none"
  doubleStrike: true,
  subScript: true,
  superScript: true,
  font: "Calibri",
  size: 24, // Half-points (24 = 12pt)
  color: "FF0000",
  highlight: "yellow", // Word highlight colors
  shading: { type: "clear", fill: "E0E0E0" },
  characterSpacing: 100,
  break: 1, // Line breaks
  tab: { type: "left", position: 1000 },
  border: {
    color: "000000",
    style: "single",
    space: 1,
    size: 6,
  },
});
```

### Paragraph Options

```ts
new Paragraph({
    text: "Simple text",        // Shorthand for single TextRun
    heading: HeadingLevel.HEADING_1,
    alignment: AlignmentType.CENTER,  // LEFT | CENTER | RIGHT | JUSTIFIED
    spacing: {
        before: 240,            // Twips
        after: 240,
        line: 360,              // Line spacing in 240ths of a line
        lineRule: "auto",       // "auto" | "exact" | "atLeast"
    },
    indent: {
        left: 720,              // Twips
        right: 720,
        firstLine: 360,
        hanging: 360,
    },
    numbering: { reference: "my-numbering", level: 0 },
    shading: { type: "clear", fill: "F0F0F0" },
    border: {
        top: { style: "single", size: 6, color: "000000", space: 1 },
        bottom: { style: "single", size: 6, color: "000000", space: 1 },
        left: { style: "single", size: 6, color: "000000", space: 1 },
        right: { style: "single", size: 6, color: "000000", space: 1 },
    },
    children: [...],            // TextRun[], ImageRun[], etc.
});
```

## Images

```ts
import { ImageRun } from "@office-open/docx";

// Basic
new ImageRun({
  type: "png",
  data: imageBuffer,
  transformation: { width: 200, height: 150 },
});

// With effects
new ImageRun({
  type: "jpg",
  data: imageBuffer,
  transformation: {
    width: 200,
    height: 150,
    rotation: 45,
    flip: { horizontal: true },
  },
  srcRect: { left: 1000, top: 1000, right: 1000, bottom: 1000 },
  blipEffects: {
    grayscale: true,
    luminance: { bright: 30, contrast: -20 },
    hsl: { hue: 0, saturation: 50, luminance: 0 },
    tint: { hue: 6000000, amount: 40 },
    duotone: { color1: { value: "002060" }, color2: { value: "D0CECE" } },
    biLevel: { threshold: 50 },
  },
});

// Floating image
new ImageRun({
  type: "png",
  data: imageBuffer,
  transformation: { width: 150, height: 150 },
  floating: {
    horizontalPosition: { offset: 720000 },
    verticalPosition: { offset: 720000 },
    wrap: { type: "square" },
  },
});

// SVG with fallback
new ImageRun({
  type: "svg",
  data: svgBuffer,
  transformation: { width: 200, height: 200 },
  fallback: { type: "png", data: pngBuffer },
});
```

Supported types: `"jpg"`, `"png"`, `"gif"`, `"bmp"`, `"svg"`, `"emf"`, `"wmf"`, `"tiff"`.

## Tables

```ts
import { Table, TableRow, TableCell, WidthType, VerticalAlign } from "@office-open/docx";

new Table({
  width: { size: 100, units: WidthType.PERCENTAGE },
  rows: [
    new TableRow({
      children: [
        new TableCell({
          width: { size: 50, units: WidthType.PERCENTAGE },
          shading: { fill: "4472C4" },
          children: [new Paragraph("Header 1")],
        }),
        new TableCell({
          children: [new Paragraph("Header 2")],
        }),
      ],
    }),
  ],
});
```

### Cell Options

- `columnSpan`, `rowSpan` — Merging cells
- `verticalAlign` — `VerticalAlign.TOP | CENTER | BOTTOM`
- `borders` — Cell-level border overrides
- `margins` — Cell padding `{ top, bottom, left, right }` in twips

## Headers & Footers

```ts
import { Header, Footer, PageNumber, NumberOfPages } from "@office-open/docx";

new Document({
    sections: [{
        headers: {
            default: new Header({
                children: [new Paragraph("Header")],
            }),
            first: new Header({
                children: [new Paragraph("First page header")],
            }),
        },
        footers: {
            default: new Footer({
                children: [
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                            new TextRun("Page "),
                            PageNumber.CURRENT,
                            new TextRun(" of "),
                            NumberOfPages.TOTAL_PAGES,
                        ],
                    }),
                ],
            }),
        },
        children: [...],
    }],
});
```

## Page Layout

```ts
new Document({
    sections: [
        // Portrait A4
        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838, orientation: "portrait" },
                    margins: {
                        top: 1440, bottom: 1440, left: 1440, right: 1440,
                        gutter: 0, header: 720, footer: 720,
                    },
                    columns: { count: 2, space: 708 },
                    borders: {
                        top: { style: "single", size: 6, color: "000000", space: 24 },
                        bottom: { style: "single", size: 6, color: "000000", space: 24 },
                        left: { style: "single", size: 6, color: "000000", space: 24 },
                        right: { style: "single", size: 6, color: "000000", space: 24 },
                    },
                },
                type: SectionType.NEXT_PAGE,
            },
            children: [...],
        },
    ],
});
```

## Math Equations

```ts
import {
  Math,
  MathRun,
  MathFraction,
  MathSubScript,
  MathSuperScript,
  MathRadical,
  MathNary,
  MathAccent,
  MathBorderBox,
  MathBox,
  MathEqArr,
  MathMatrix,
  MathGroupChr,
  MathPhant,
} from "@office-open/docx";

// Fraction
new MathFraction({ numerator: "a", denominator: "b" });

// Superscript: x²
new MathSuperScript({ children: [new MathRun("x")], superScript: [new MathRun("2")] });

// Square root
new MathRadical({ children: [new MathRun("a + b")] });

// Summation
new MathNary({ operator: "∑", children: [new MathRun("i=1..n")] });

// Border box with hidden borders
new MathBorderBox({
  children: [new MathRun("a")],
  properties: { hideTop: true, hideBottom: true, strikeHorizontal: true },
});

// Equation array
new MathEqArr({ rows: [[new MathRun("x + y = 1")], [new MathRun("2x - y = 3")]] });

// Matrix
new MathMatrix({
  rows: [
    [new MathRun("1"), new MathRun("0")],
    [new MathRun("0"), new MathRun("1")],
  ],
});
```

## Links & Bookmarks

```ts
import {
  ExternalHyperlink,
  InternalHyperlink,
  BookmarkStart,
  BookmarkEnd,
} from "@office-open/docx";

new ExternalHyperlink({
  children: [new TextRun("Click here")],
  link: "https://example.com",
});

new Paragraph({
  children: [
    new BookmarkStart("my-bookmark"),
    new TextRun("Target text"),
    new BookmarkEnd("my-bookmark"),
  ],
});
```

## Styles & Numbering

```ts
new Document({
    styles: {
        default: {
            document: {
                run: { font: "Calibri", size: 24 },
            },
        },
        paragraphStyles: [
            {
                id: "myStyle",
                name: "My Custom Style",
                basedOn: "Normal",
                run: { bold: true, color: "0000FF" },
            },
        ],
    },
    numbering: {
        config: [
            {
                reference: "my-list",
                levels: [
                    { level: 0, format: "decimal", text: "%1.", alignment: "left" },
                ],
            },
        ],
    },
    sections: [...],
});
```

## Shapes (WpsShapeRun)

```ts
import { WpsShapeRun } from "@office-open/docx";

new WpsShapeRun({
  children: [new Paragraph({ children: [new TextRun("Shape text")] })],
  customGeometry: {
    pathList: [
      {
        w: 100000,
        h: 100000,
        commands: [
          { command: "moveTo", point: { x: "50000", y: "0" } },
          { command: "lineTo", point: { x: "100000", y: "100000" } },
          { command: "lineTo", point: { x: "0", y: "100000" } },
          { command: "close" },
        ],
      },
    ],
  },
  fill: "4472C4",
  transformation: { height: 150, width: 200 },
  type: "wps",
});
```

## Comments & Revisions

```ts
import { Comment, CommentReference } from "@office-open/docx";

new Document({
  comments: {
    children: [
      new Comment({
        id: 0,
        author: "User",
        children: [new Paragraph("This is a comment")],
      }),
    ],
  },
  sections: [
    {
      children: [
        new Paragraph({
          children: [new TextRun("Some text"), new CommentReference(0)],
        }),
      ],
    },
  ],
});
```
