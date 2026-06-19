# DOCX API Reference

Complete API reference for `@office-open/docx`. All examples show the options JSON structure. Pass these options objects to `generateDocument()`.

## Document Structure

```
DocumentOptions
├── sections: SectionOptions[]
│   ├── properties: SectionPropertiesOptionsBase
│   │   ├── page: { size, margins, columns, borders }
│   │   └── type: "nextPage" | "continuous" | "evenPage" | "oddPage"
│   ├── headers: { default, first, even, odd }
│   ├── footers: { default, first, even, odd }
│   └── children: (ParagraphOptions | TableOptions | ImageOptions | ...)[]
```

## Text Formatting

### TextRun Options

```json
{
  "text": "Hello",
  "bold": true,
  "italic": true,
  "underline": { "type": "single", "color": "FF0000" },
  "strike": "single",
  "doubleStrike": true,
  "subScript": true,
  "superScript": true,
  "font": "Calibri",
  "size": 12,
  "color": "FF0000",
  "highlight": "yellow",
  "shading": { "type": "clear", "fill": "E0E0E0" },
  "characterSpacing": 100,
  "break": 1,
  "tab": { "type": "left", "position": 1000 },
  "border": {
    "color": "000000",
    "style": "single",
    "space": 1,
    "size": 6
  }
}
```

| Property           | Type                             | Description                                                                          |
| ------------------ | -------------------------------- | ------------------------------------------------------------------------------------ |
| `text`             | `string`                         | Plain text content                                                                   |
| `bold`             | `boolean`                        | Bold formatting                                                                      |
| `italic`           | `boolean`                        | Italic formatting                                                                    |
| `underline`        | `{ type, color? }`               | Underline style. Types: `"single"`, `"double"`, `"wave"`, `"dash"`, `"dotted"`, etc. |
| `strike`           | `"single" \| "double" \| "none"` | Strikethrough                                                                        |
| `size`             | `number`                         | Font size in points (12 = 12pt)                                                      |
| `color`            | `string`                         | Hex color without `#`                                                                |
| `font`             | `string`                         | Font family name                                                                     |
| `highlight`        | `string`                         | Word highlight color name                                                            |
| `shading`          | `{ type, fill }`                 | Background shading                                                                   |
| `characterSpacing` | `number`                         | Character spacing in twips                                                           |
| `break`            | `number`                         | Number of line breaks                                                                |
| `tab`              | `{ type, position }`             | Tab stop                                                                             |
| `children`         | `array`                          | Mixed content: strings, PageNumber tokens, Math elements, CommentReference, etc.     |

### Paragraph Options

```json
{
  "text": "Simple text",
  "heading": "Heading1",
  "alignment": "center",
  "spacing": {
    "before": 240,
    "after": 240,
    "line": 360,
    "lineRule": "auto"
  },
  "indent": {
    "left": 720,
    "right": 720,
    "firstLine": 360,
    "hanging": 360
  },
  "numbering": { "reference": "my-numbering", "level": 0 },
  "shading": { "type": "clear", "fill": "F0F0F0" },
  "border": {
    "top": { "style": "single", "size": 6, "color": "000000", "space": 1 },
    "bottom": { "style": "single", "size": 6, "color": "000000", "space": 1 },
    "left": { "style": "single", "size": 6, "color": "000000", "space": 1 },
    "right": { "style": "single", "size": 6, "color": "000000", "space": 1 }
  },
  "children": []
}
```

| Property    | Type                                      | Description                                                        |
| ----------- | ----------------------------------------- | ------------------------------------------------------------------ |
| `text`      | `string`                                  | Shorthand for single TextRun child                                 |
| `heading`   | `string`                                  | `"Heading1"` through `"Heading6"`                                  |
| `alignment` | `string`                                  | `"start"` \| `"center"` \| `"end"` \| `"both"` \| `"distribute"`   |
| `spacing`   | `{ before?, after?, line?, lineRule? }`   | Spacing in twips. `lineRule`: `"auto"` \| `"exact"` \| `"atLeast"` |
| `indent`    | `{ left?, right?, firstLine?, hanging? }` | Indentation in twips                                               |
| `numbering` | `{ reference, level }`                    | List reference and level                                           |
| `children`  | `array`                                   | TextRun, ImageRun, Math, Bookmark, etc.                            |

## Images

### Basic Image

```json
{
  "type": "png",
  "data": "<Uint8Array>",
  "transformation": { "width": 200, "height": 150 }
}
```

### Image with Effects

```json
{
  "type": "jpg",
  "data": "<Uint8Array>",
  "transformation": {
    "width": 200,
    "height": 150,
    "rotation": 45,
    "flip": { "horizontal": true }
  },
  "srcRect": { "left": 1000, "top": 1000, "right": 1000, "bottom": 1000 },
  "blipEffects": {
    "grayscale": true,
    "luminance": { "bright": 30, "contrast": -20 },
    "hsl": { "hue": 0, "saturation": 50, "luminance": 0 },
    "tint": { "hue": 6000000, "amount": 40 },
    "duotone": { "color1": { "value": "002060" }, "color2": { "value": "D0CECE" } },
    "biLevel": { "threshold": 50 }
  }
}
```

### Floating Image

```json
{
  "type": "png",
  "data": "<Uint8Array>",
  "transformation": { "width": 150, "height": 150 },
  "floating": {
    "horizontalPosition": { "offset": 720000 },
    "verticalPosition": { "offset": 720000 },
    "wrap": { "type": "square" }
  }
}
```

### SVG with Fallback

```json
{
  "type": "svg",
  "data": "<Uint8Array>",
  "transformation": { "width": 200, "height": 200 },
  "fallback": { "type": "png", "data": "<Uint8Array>" }
}
```

Supported types: `"jpg"`, `"png"`, `"gif"`, `"bmp"`, `"svg"`, `"emf"`, `"wmf"`, `"tiff"`.

## Tables

### Table Options

```json
{
  "width": { "size": 100, "units": "pct" },
  "rows": [
    {
      "cells": [
        {
          "width": { "size": 50, "units": "pct" },
          "shading": { "fill": "4472C4" },
          "children": [{ "text": "Header 1" }]
        },
        {
          "children": [{ "text": "Header 2" }]
        }
      ]
    }
  ]
}
```

| `width.units` values | Description                   |
| -------------------- | ----------------------------- |
| `"auto"`             | Automatic width               |
| `"dxa"`              | Twips (twentieths of a point) |
| `"pct"`              | Percentage (50 = 50%)         |

### Cell Options

| Property        | Type                            | Description                 |
| --------------- | ------------------------------- | --------------------------- |
| `columnSpan`    | `number`                        | Horizontal merge            |
| `rowSpan`       | `number`                        | Vertical merge              |
| `verticalAlign` | `"top" \| "center" \| "bottom"` | Vertical alignment          |
| `borders`       | `object`                        | Cell-level border overrides |
| `margins`       | `{ top, bottom, left, right }`  | Cell padding in twips       |

## Headers & Footers

```json
{
  "headers": {
    "default": {
      "children": [{ "text": "Header" }]
    },
    "first": {
      "children": [{ "text": "First page header" }]
    }
  },
  "footers": {
    "default": {
      "children": [
        {
          "alignment": "center",
          "children": ["Page ", "CURRENT", " of ", "TOTAL_PAGES"]
        }
      ]
    }
  }
}
```

Page number tokens (used as string values in TextRun children):

- `"CURRENT"` — Current page number
- `"TOTAL_PAGES"` — Total page count
- `"TOTAL_PAGES_IN_SECTION"` — Section page count
- `"SECTION"` — Current section page

## Page Layout

```json
{
  "sections": [
    {
      "properties": {
        "page": {
          "size": { "width": 11906, "height": 16838, "orientation": "portrait" },
          "margins": {
            "top": 1440,
            "bottom": 1440,
            "left": 1440,
            "right": 1440,
            "gutter": 0,
            "header": 720,
            "footer": 720
          },
          "columns": { "count": 2, "space": 708 },
          "borders": {
            "top": { "style": "single", "size": 6, "color": "000000", "space": 24 },
            "bottom": { "style": "single", "size": 6, "color": "000000", "space": 24 },
            "left": { "style": "single", "size": 6, "color": "000000", "space": 24 },
            "right": { "style": "single", "size": 6, "color": "000000", "space": 24 }
          }
        },
        "type": "nextPage"
      },
      "children": []
    }
  ]
}
```

| `type` values  | Description                      |
| -------------- | -------------------------------- |
| `"nextPage"`   | Section begins on the next page  |
| `"continuous"` | Section continues on same page   |
| `"evenPage"`   | Section begins on next even page |
| `"oddPage"`    | Section begins on next odd page  |

## Math Equations

### Available Math Elements

| Element                 | Key Properties                                                                |
| ----------------------- | ----------------------------------------------------------------------------- |
| `MathFraction`          | `{ numerator: [...], denominator: [...] }` — arrays of MathRun                |
| `MathSuperScript`       | `{ children: [...], superScript: [...] }`                                     |
| `MathSubScript`         | `{ children: [...], subScript: [...] }`                                       |
| `MathSubSuperScript`    | `{ children: [...], subScript: [...], superScript: [...] }`                   |
| `MathPreSubSuperScript` | `{ children: [...], subScript: [...], superScript: [...] }`                   |
| `MathRadical`           | `{ children: [...], degree?: [...] }`                                         |
| `MathSum`               | `{ children: [...], subScript?: [...], superScript?: [...] }`                 |
| `MathIntegral`          | `{ children: [...], subScript?: [...], superScript?: [...] }`                 |
| `MathFunction`          | `{ children: [...], name: [...] }`                                            |
| `MathLimitUpper`        | `{ children: [...], limit: [...] }`                                           |
| `MathLimitLower`        | `{ children: [...], limit: [...] }`                                           |
| `MathRoundBrackets`     | `{ children: [...] }`                                                         |
| `MathSquareBrackets`    | `{ children: [...] }`                                                         |
| `MathCurlyBrackets`     | `{ children: [...] }`                                                         |
| `MathAngledBrackets`    | `{ children: [...] }`                                                         |
| `MathBorderBox`         | `{ children: [...], properties?: { hideTop, hideBottom, strikeHorizontal } }` |
| `MathEqArr`             | `{ rows: [[...], [...]] }`                                                    |
| `MathMatrix`            | `{ rows: [[...], [...]] }`                                                    |
| `MathGroupChr`          | `{ children: [...], properties?: { chr, pos, vertJc } }`                      |
| `MathPhant`             | `{ children: [...], properties?: { zeroAsc, zeroDesc } }`                     |
| `MathBox`               | `{ children: [...], properties?: { opEmu } }`                                 |
| `MathRun`               | `{ text: "content" }` or `"string content"`                                   |
| `createMathAccent`      | `{ children: [...], accentCharacter?: string }` (factory function)            |

### Examples

```json
// Fraction: a/b
{ "numerator": ["a"], "denominator": ["b"] }

// Fraction with type: "skw" (skewed), "lin" (linear), "noBar" (no fraction bar)
{ "numerator": ["a"], "denominator": ["b"], "fractionType": "skw" }

// Superscript: x²
{ "children": ["x"], "superScript": ["2"] }

// Square root: √(a+b)
{ "children": ["a + b"] }

// Nth root: ³√(a+b)
{ "children": ["a + b"], "degree": ["3"] }

// Sum with limits: Σᵢˡ⁰
{ "children": ["test"], "subScript": ["i"], "superScript": ["10"] }

// Function: sin(100)
{ "children": ["100"], "name": ["sin"] }

// Border box with hidden borders
{
  "children": ["b"],
  "properties": { "hideTop": true, "hideBottom": true }
}

// Equation array
{ "rows": [["x + y = 1"], ["2x - y = 3"]] }

// Matrix
{ "rows": [["1", "0"], ["0", "1"]] }

// Accent (via factory function)
{ "children": ["x"] }
{ "children": ["y"], "accentCharacter": "\u0303" }
```

**Note**: All `numerator`, `denominator`, `children`, `superScript`, `subScript`, etc. accept arrays of MathRun elements. String values are auto-wrapped as MathRun.

## Links & Bookmarks

### External Hyperlink

```json
{
  "children": [{ "text": "Click here" }],
  "link": "https://example.com"
}
```

### Bookmark

```json
{
  "id": "myAnchorId",
  "children": [{ "text": "Lorem Ipsum" }]
}
```

### Internal Hyperlink (links to a Bookmark)

```json
{
  "children": [{ "text": "Go to section", "bold": true }],
  "anchor": "myAnchorId"
}
```

### Page Reference (shows page number of a bookmark)

```json
"myAnchorId"
```

Used in TextRun children as a string — displays the page number where the referenced bookmark is located.

## Styles & Numbering

```json
{
  "styles": {
    "default": {
      "document": {
        "run": { "font": "Calibri", "size": 12 }
      }
    },
    "paragraphStyles": [
      {
        "id": "myStyle",
        "name": "My Custom Style",
        "basedOn": "Normal",
        "run": { "bold": true, "color": "0000FF" }
      }
    ]
  },
  "numbering": {
    "config": [
      {
        "reference": "my-list",
        "levels": [
          {
            "level": 0,
            "format": "decimal",
            "text": "%1.",
            "alignment": "left"
          }
        ]
      }
    ]
  },
  "sections": []
}
```

## Shapes (WpsShapeRun)

```json
{
  "children": [{ "alignment": "center", "children": [{ "text": "Shape text" }] }],
  "customGeometry": {
    "pathList": [
      {
        "w": 100000,
        "h": 100000,
        "commands": [
          { "command": "moveTo", "point": { "x": "50000", "y": "0" } },
          { "command": "lineTo", "point": { "x": "100000", "y": "100000" } },
          { "command": "lineTo", "point": { "x": "0", "y": "100000" } },
          { "command": "close" }
        ]
      }
    ]
  },
  "fill": "4472C4",
  "outline": { "color": { "value": "C00000" }, "type": "solidFill", "width": 12700 },
  "transformation": { "height": 150, "width": 200 },
  "type": "wps"
}
```

**Note**: `type: "wps"` is required. `children` contains paragraph options. Command types: `"moveTo"`, `"lineTo"`, `"arcTo"` (with `heightRadius`, `widthRadius`, `startAngle`, `sweepAngle`), `"close"`.

## Comments & Revisions

```json
{
  "comments": {
    "children": [
      {
        "id": 0,
        "author": "User",
        "date": "<Date object>",
        "children": [{ "text": "This is a comment" }]
      }
    ]
  },
  "sections": [
    {
      "children": [
        {
          "children": [
            "Some text",
            { "commentRangeStart": 0 },
            { "text": "commented text" },
            { "commentRangeEnd": 0 },
            { "children": [{ "commentReference": 0 }] }
          ]
        }
      ]
    }
  ]
}
```

**Note**: Use `{ "commentRangeStart": 0 }`, `{ "commentRangeEnd": 0 }`, and `{ "children": [{ "commentReference": 0 }] }` as separate elements in the children array. `commentReference` goes inside a TextRun's children.

## Patching

### patchDocument

Modify an existing `.docx` template by replacing placeholders:

```json
// Replace a paragraph placeholder
{
  "type": "paragraph",
  "children": [{ "text": "John Doe" }]
}

// Replace with block-level content
{
  "type": "document",
  "children": [
    { "children": [{ "text": "First paragraph" }] },
    { "children": [{ "text": "Second paragraph" }] }
  ]
}
```

```ts
import { patchDocument } from "@office-open/docx";

const result = await patchDocument({
  outputType: "nodebuffer",
  data: templateBuffer,
  placeholders: {
    name: { type: "paragraph", children: [{ text: "John Doe" }] },
    content: {
      type: "document",
      children: [{ children: ["First"] }, { children: ["Second"] }],
    },
  },
  placeholderDelimiters: { start: "{{", end: "}}" },
  keepOriginalStyles: true,
  recursive: true,
});
```

Use `findReplace` for literal text (no delimiters) and `coreProperties` to override `docProps/core.xml`.

| Option                  | Type                             | Default      | Description                                 |
| ----------------------- | -------------------------------- | ------------ | ------------------------------------------- |
| `outputType`            | `string`                         | —            | Output format                               |
| `data`                  | `Buffer \| Uint8Array \| ...`    | —            | Input .docx file                            |
| `placeholders`          | `Record<string, Patch>`          | —            | Delimiter-wrapped placeholder → patch       |
| `findReplace`           | `Record<string, Patch>`          | —            | Literal find string → patch                 |
| `coreProperties`        | `Partial<CorePropertiesOptions>` | —            | Core metadata override                      |
| `keepOriginalStyles`    | `boolean`                        | `true`       | Preserve original run formatting properties |
| `placeholderDelimiters` | `{ start, end }`                 | `{ {{, }} }` | Custom delimiters                           |
| `recursive`             | `boolean`                        | `true`       | Replace all occurrences                     |

| Patch type    | Value         | Description                  |
| ------------- | ------------- | ---------------------------- |
| `"paragraph"` | `"paragraph"` | Inline run-level replacement |
| `"document"`  | `"document"`  | Block-level replacement      |

### patchDetector

Scan a template to discover all placeholder keys:

```ts
import { patchDetector } from "@office-open/docx";

const placeholders = await patchDetector({ data: templateBuffer });
// ["name", "title", "content", ...]
```
