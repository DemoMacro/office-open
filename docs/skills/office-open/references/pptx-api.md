# PPTX API Reference

Complete API reference for `@office-open/pptx`. All examples show the options JSON structure.

## Presentation Structure

```json
{
  "title": "string",
  "creator": "string",
  "subject": "string",
  "description": "string",
  "slides": [
    {
      "background": {},
      "transition": {},
      "children": []
    }
  ]
}
```

Slide children types: shape, table, image, group, line, connector.

## Shapes

### Shape Options

```json
{
  "id": 1,
  "name": "My Shape",
  "x": 50,
  "y": 100,
  "width": 300,
  "height": 80,
  "text": "Hello World",
  "fill": "4472C4"
}
```

| Property    | Type                    | Description                                           |
| ----------- | ----------------------- | ----------------------------------------------------- |
| `id`        | `number`                | Optional, auto-assigned                               |
| `name`      | `string`                | Optional display name                                 |
| `x`         | `number`                | Position X (pixels)                                   |
| `y`         | `number`                | Position Y (pixels)                                   |
| `width`     | `number`                | Width (pixels)                                        |
| `height`    | `number`                | Height (pixels)                                       |
| `text`      | `string`                | Simple text content (shorthand)                       |
| `textBody`  | `object`                | Rich text body (see below)                            |
| `fill`      | `string \| FillOptions` | Fill style. String `"4472C4"` = solid color shorthand |
| `outline`   | `OutlineOptions`        | Outline/border                                        |
| `effects`   | `EffectsOptions`        | Visual effects                                        |
| `geometry`  | `string`                | Preset geometry                                       |
| `rotation`  | `number`                | Rotation in degrees                                   |
| `flipV`     | `boolean`               | Flip vertically                                       |
| `flipH`     | `boolean`               | Flip horizontally                                     |
| `animation` | `AnimationOptions`      | Animation settings                                    |

### Shape with Rich Text (textBody)

```json
{
  "x": 50,
  "y": 100,
  "width": 400,
  "height": 120,
  "textBody": {
    "children": [
      {
        "children": [{ "text": "Bold title", "bold": true, "fontSize": 24, "color": "FFFFFF" }],
        "alignment": "center"
      },
      {
        "children": [{ "text": "Subtitle", "fontSize": 16, "color": "CCCCCC" }]
      }
    ]
  },
  "fill": { "type": "solid", "color": "2E74B5" }
}
```

### textBody Options

```json
{
  "text": "Content",
  "anchor": "center",
  "vertical": "vert",
  "autoFit": "normal",
  "wrap": "square",
  "margins": {
    "top": 45720,
    "bottom": 45720,
    "left": 91440,
    "right": 91440
  },
  "columns": 2,
  "columnSpacing": 45720
}
```

| Property        | Type     | Values                                                                                                                    |
| --------------- | -------- | ------------------------------------------------------------------------------------------------------------------------- |
| `anchor`        | `string` | `"top"` \| `"center"` \| `"bottom"` \| `"justify"` \| `"distribute"`                                                      |
| `vertical`      | `string` | `"horz"` (default) \| `"vert"` \| `"vert270"` \| `"wordArtVert"` \| `"eaVert"` \| `"mongolianVert"` \| `"wordArtVertRtl"` |
| `autoFit`       | `string` | `"none"` \| `"normal"` \| `"shape"`                                                                                       |
| `wrap`          | `string` | `"square"` \| `"none"`                                                                                                    |
| `margins`       | `object` | EMU values                                                                                                                |
| `columns`       | `number` | Number of text columns                                                                                                    |
| `columnSpacing` | `number` | Spacing between columns (EMU)                                                                                             |

### Geometry Presets

```json
{ "geometry": "rect" }
```

Common values: `"rect"`, `"roundRect"`, `"ellipse"`, `"triangle"`, `"diamond"`, `"hexagon"`, `"star5"`, `"star6"`, `"arrow"`, `"chevron"`, `"cloud"`, `"heart"`, `"lightningBolt"`, `"moon"`, `"sun"`.

## Fills

### Solid Fill

```json
{ "type": "solid", "color": "4472C4" }
```

String shorthand: `fill: "4472C4"` is equivalent to the above.

### Gradient Fill

```json
{
  "type": "gradient",
  "stops": [
    { "position": 0, "color": "0D47A1" },
    { "position": 50, "color": "1976D2" },
    { "position": 100, "color": "42A5F5" }
  ],
  "angle": 45
}
```

| Property | Type     | Description                           |
| -------- | -------- | ------------------------------------- |
| `stops`  | `array`  | Array of `{ position: 0-100, color }` |
| `angle`  | `number` | Angle in degrees                      |
| `path`   | `string` | `"circle"` for radial gradient        |

### Image Fill

```json
{
  "type": "blip",
  "data": "<Uint8Array>",
  "imageType": "png"
}
```

### No Fill

```json
{ "type": "noFill" }
```

## Outline

```json
{
  "color": "1565C0",
  "width": 2,
  "dashStyle": "solid",
  "capStyle": "flat",
  "compound": "single"
}
```

| Property    | Type     | Values                                                        |
| ----------- | -------- | ------------------------------------------------------------- |
| `color`     | `string` | Hex color without `#`                                         |
| `width`     | `number` | Width in points                                               |
| `dashStyle` | `string` | `"solid"` \| `"dash"` \| `"dashDot"` \| `"dot"` \| `"lgDash"` |
| `capStyle`  | `string` | `"flat"` \| `"round"` \| `"square"`                           |
| `compound`  | `string` | `"single"` \| `"double"` \| `"thickThin"`                     |

## Effects

### Outer Shadow

```json
{
  "outerShadow": {
    "blur": 50800,
    "distance": 38100,
    "direction": 5400000,
    "color": "000000",
    "alpha": 50,
    "rotateWithShape": true
  }
}
```

### Inner Shadow

```json
{
  "innerShadow": {
    "blur": 40000,
    "distance": 30000,
    "direction": 5400000,
    "color": "000000",
    "alpha": 40
  }
}
```

### Glow

```json
{
  "glow": {
    "radius": 152400,
    "color": "92D050",
    "alpha": 60
  }
}
```

### Reflection

```json
{
  "reflection": {
    "blurRadius": 6350,
    "distance": 38100,
    "direction": 5400000,
    "startAlpha": 90,
    "endAlpha": 0,
    "startPosition": 0,
    "endPosition": 100
  }
}
```

### Soft Edge

```json
{
  "softEdge": { "radius": 50800 }
}
```

### 3D Rotation

```json
{
  "rotation3D": {
    "x": 20,
    "y": 30,
    "z": 10,
    "perspective": 500
  }
}
```

| Property      | Type     | Description         |
| ------------- | -------- | ------------------- |
| `x`, `y`, `z` | `number` | Rotation in degrees |
| `perspective` | `number` | Perspective value   |

### Extrusion & Bevel

```json
{
  "rotation3D": { "x": 25, "y": 15 },
  "extrusionH": 50000,
  "material": "plastic",
  "bevelTop": { "width": 8, "height": 8 },
  "bevelBottom": { "width": 5, "height": 5 },
  "lighting": "threePt"
}
```

| Property      | Type                | Values                                                                           |
| ------------- | ------------------- | -------------------------------------------------------------------------------- |
| `extrusionH`  | `number`            | Extrusion height (EMU)                                                           |
| `material`    | `string`            | `"plastic"` \| `"metal"` \| `"matte"` \| `"warmMatte"` \| `"flat"` \| `"powder"` |
| `bevelTop`    | `{ width, height }` | Bevel dimensions in points                                                       |
| `bevelBottom` | `{ width, height }` | Bevel dimensions in points                                                       |
| `lighting`    | `string`            | `"flat"` \| `"threePt"` \| `"balanced"` \| `"soft"` \| `"harsh"`                 |

## Tables

```json
{
  "x": 50,
  "y": 200,
  "width": 600,
  "height": 250,
  "rows": [
    {
      "cells": [
        { "text": "Name", "fill": "4472C4", "bold": true, "color": "FFFFFF" },
        { "text": "Age" },
        { "text": "City" }
      ]
    },
    {
      "cells": [{ "text": "Alice" }, { "text": "30" }, { "text": "Beijing" }]
    }
  ],
  "columnWidths": [2000000, 1500000, 2500000],
  "firstRow": true,
  "bandRow": true
}
```

| Property       | Type       | Description                                            |
| -------------- | ---------- | ------------------------------------------------------ |
| `rows`         | `array`    | Each row has `cells` array and optional `height` (EMU) |
| `columnWidths` | `number[]` | Column widths in EMU                                   |
| `firstRow`     | `boolean`  | Apply header styling to first row                      |
| `bandRow`      | `boolean`  | Banded row styling                                     |

### Cell Options

| Property        | Type                            | Description        |
| --------------- | ------------------------------- | ------------------ |
| `text`          | `string`                        | Cell text          |
| `fill`          | `string`                        | Background color   |
| `bold`          | `boolean`                       | Bold text          |
| `color`         | `string`                        | Text color         |
| `verticalAlign` | `"top" \| "center" \| "bottom"` | Vertical alignment |
| `margins`       | `{ left, right, top, bottom }`  | Cell padding (EMU) |
| `columnSpan`    | `number`                        | Column merge       |
| `rowSpan`       | `number`                        | Row merge          |

## Images

```json
{
  "x": 100,
  "y": 100,
  "width": 300,
  "height": 200,
  "data": "<Uint8Array>",
  "imageType": "png"
}
```

## Charts

Charts are created via `@office-open/core` and embedded in slides.

## Lines & Connectors

### Simple Line

```json
{
  "x1": 50,
  "y1": 120,
  "x2": 800,
  "y2": 120,
  "outline": { "color": "4472C4", "width": 2 }
}
```

### Connector (line with arrowheads)

```json
{
  "x1": 50,
  "y1": 130,
  "x2": 400,
  "y2": 130,
  "outline": { "color": "4472C4", "width": 2 },
  "endArrowhead": "triangle",
  "beginArrowhead": "triangle"
}
```

| Property          | Type     | Values                                                             |
| ----------------- | -------- | ------------------------------------------------------------------ |
| `beginArrowhead`  | `string` | `"triangle"` \| `"stealth"` \| `"diamond"` \| `"oval"` \| `"open"` |
| `endArrowhead`    | `string` | Same as above                                                      |
| `arrowheadWidth`  | `string` | `"sm"` \| `"med"` \| `"lg"`                                        |
| `arrowheadLength` | `string` | `"sm"` \| `"med"` \| `"lg"`                                        |

## Group Shapes

```json
{
  "x": 50,
  "y": 50,
  "width": 500,
  "height": 300,
  "rotation": 10,
  "children": [
    { "x": 0, "y": 0, "width": 240, "height": 140, "textBody": { "text": "A" }, "fill": "4472C4" },
    { "x": 260, "y": 0, "width": 240, "height": 140, "textBody": { "text": "B" }, "fill": "ED7D31" }
  ]
}
```

**Note**: Children coordinates are relative to the group. Supports nested groups.

## Slide Background

```json
{
  "slides": [
    {
      "background": { "type": "solid", "color": "FFFFFF" },
      "children": []
    }
  ]
}
```

## Transitions

Transitions are set on slide options (not on shapes):

```json
{
  "slides": [
    {
      "transition": {
        "type": "fade",
        "speed": "med"
      },
      "children": []
    }
  ]
}
```

```json
{
  "slides": [
    {
      "transition": {
        "type": "push",
        "direction": "right",
        "speed": "slow"
      },
      "children": []
    }
  ]
}
```

| Property           | Type                        | Description                                                                                                                    |
| ------------------ | --------------------------- | ------------------------------------------------------------------------------------------------------------------------------ |
| `type`             | `string`                    | See values below                                                                                                               |
| `speed`            | `"slow" \| "med" \| "fast"` | Transition speed                                                                                                               |
| `advanceOnClick`   | `boolean`                   | Advance on click                                                                                                               |
| `advanceAfterTime` | `number`                    | Auto-advance after ms                                                                                                          |
| `direction`        | `string`                    | `"left"` \| `"up"` \| `"right"` \| `"down"` \| `"leftUp"` \| `"rightUp"` \| `"leftDown"` \| `"rightDown"` \| `"in"` \| `"out"` |
| `orient`           | `"horz" \| "vert"`          | Orientation (for split, blinds, etc.)                                                                                          |
| `spokes`           | `number`                    | Number of spokes for wheel transition                                                                                          |

Transition types: `"fade"`, `"push"`, `"wipe"`, `"split"`, `"blinds"`, `"checker"`, `"comb"`, `"randomBar"`, `"cover"`, `"pull"`, `"strips"`, `"wheel"`, `"zoom"`, `"circle"`, `"dissolve"`, `"diamond"`, `"newsflash"`, `"plus"`, `"wedge"`, `"random"`, `"cut"`.

## Animations

Animations are set on Shape options:

### Entrance/Exit Animations

```json
{
  "animation": { "type": "fade", "duration": 800 }
}
```

```json
{
  "animation": { "type": "appear" }
}
```

```json
{
  "animation": { "type": "fly", "direction": "left", "duration": 600 }
}
```

```json
{
  "animation": { "type": "zoom" }
}
```

```json
{
  "animation": { "type": "split", "direction": "horizontal" }
}
```

### Exit Animations

Add `class: "exit"`:

```json
{
  "animation": { "type": "fade", "class": "exit", "duration": 800 }
}
```

```json
{
  "animation": { "type": "fly", "class": "exit", "direction": "right", "duration": 600 }
}
```

### Emphasis Animations

```json
{
  "animation": { "class": "emph", "emphasisType": "growShrink", "duration": 800 }
}
```

```json
{
  "animation": { "class": "emph", "emphasisType": "spin", "duration": 1000 }
}
```

```json
{
  "animation": {
    "class": "emph",
    "emphasisType": "colorChange",
    "color": "FF0000",
    "duration": 800
  }
}
```

`emphasisType` values: `"growShrink"`, `"spin"`, `"growWithTurn"`, `"colorWave"`, `"colorChange"`, `"blink"`, `"transparency"`, `"boldFlash"`, `"wave"`, `"pulse"`.

### Path Animations

```json
{
  "animation": { "pathType": "line", "duration": 1000 }
}
```

```json
{
  "animation": { "pathType": "circle", "duration": 1500 }
}
```

```json
{
  "animation": {
    "pathType": "customPath",
    "path": "M 0 0 L 100 0 L 100 100 L 0 100 Z",
    "duration": 1200
  }
}
```

`pathType` values: `"customPath"`, `"arc"`, `"bounce"`, `"circle"`, `"curve"`, `"figureEight"`, `"line"`, `"loop"`.

### Animation Options Summary

| Property       | Type      | Values                                                                                                                        | Description                                   |
| -------------- | --------- | ----------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------- |
| `type`         | `string`  | `"appear"` \| `"fade"` \| `"fly"` \| `"wipe"` \| `"dissolve"` \| `"zoom"` \| `"cover"` \| `"push"` \| `"strips"` \| `"split"` | Entrance/exit type                            |
| `trigger`      | `string`  | `"onClick"` \| `"withPrevious"` \| `"afterPrevious"`                                                                          | How animation triggers (default: `"onClick"`) |
| `duration`     | `number`  | —                                                                                                                             | Duration in ms                                |
| `delay`        | `number`  | —                                                                                                                             | Delay before start in ms                      |
| `direction`    | `string`  | `"left"` \| `"right"` \| `"up"` \| `"down"` \| `"horizontal"` \| `"vertical"`                                                 | Animation direction                           |
| `class`        | `string`  | `"entr"` \| `"exit"` \| `"emph"`                                                                                              | Animation category                            |
| `emphasisType` | `string`  | See above                                                                                                                     | Required when `class: "emph"`                 |
| `pathType`     | `string`  | See above                                                                                                                     | Path animation type                           |
| `path`         | `string`  | SVG path syntax                                                                                                               | Custom path definition                        |
| `color`        | `string`  | Hex color                                                                                                                     | For colorChange emphasis                      |
| `speed`        | `number`  | —                                                                                                                             | Speed multiplier                              |
| `repeatCount`  | `number`  | —                                                                                                                             | Number of repetitions                         |
| `autoReverse`  | `boolean` | —                                                                                                                             | Auto-reverse animation                        |

## Parsing

```json
// parsePresentation accepts Uint8Array, ArrayBuffer, DataView, number[], base64 string
parsePresentation(readFileSync("input.pptx"))
// Returns PresentationOptions — pass to new Presentation(opts) to modify and re-export
```

## Patching

### patchPresentation

Modify an existing `.pptx` template by replacing placeholders:

```json
{
  "type": "paragraph",
  "children": [{ "text": "Updated Title", "bold": true, "fontSize": 24 }]
}
```

```ts
import { patchPresentation, PatchType, TextRun } from "@office-open/pptx";

const result = await patchPresentation({
  outputType: "nodebuffer",
  data: templateBuffer,
  patches: {
    title: { type: PatchType.PARAGRAPH, children: [new TextRun({ text: "Updated", bold: true })] },
  },
  placeholderDelimiters: { start: "{{", end: "}}" },
});
```

| Option                  | Type                          | Default      | Description                                 |
| ----------------------- | ----------------------------- | ------------ | ------------------------------------------- |
| `outputType`            | `string`                      | —            | Output format                               |
| `data`                  | `Buffer \| Uint8Array \| ...` | —            | Input .pptx file                            |
| `patches`               | `Record<string, IPatch>`      | —            | Placeholder name → patch content map        |
| `keepOriginalStyles`    | `boolean`                     | `true`       | Preserve original run formatting properties |
| `placeholderDelimiters` | `{ start, end }`              | `{ {{, }} }` | Custom delimiters                           |

| PatchType             | Value         | Description                  |
| --------------------- | ------------- | ---------------------------- |
| `PatchType.PARAGRAPH` | `"paragraph"` | Inline run-level replacement |
