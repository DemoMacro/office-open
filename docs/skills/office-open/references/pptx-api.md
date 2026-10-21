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

Slide children types: shape, picture, table, chart, group, line, connector, video, audio, smartart, ole.

## Shapes

### Shape Options

```json
{
  "id": 1,
  "name": "My Shape",
  "x": "2cm",
  "y": "4cm",
  "width": "12cm",
  "height": "3cm",
  "textBody": { "text": "Hello World" },
  "fill": "4472C4"
}
```

| Property    | Type                         | Description                                           |
| ----------- | ---------------------------- | ----------------------------------------------------- |
| `id`        | `number`                     | Optional, auto-assigned                               |
| `name`      | `string`                     | Optional display name                                 |
| `x`         | `number \| UniversalMeasure` | Position X (EMU or unit string)                       |
| `y`         | `number \| UniversalMeasure` | Position Y (EMU or unit string)                       |
| `width`     | `number \| UniversalMeasure` | Width (EMU or unit string)                            |
| `height`    | `number \| UniversalMeasure` | Height (EMU or unit string)                           |
| `textBody`  | `object`                     | Text body (see below)                                 |
| `fill`      | `string \| FillOptions`      | Fill style. String `"4472C4"` = solid color shorthand |
| `outline`   | `OutlineOptions`             | Outline/border                                        |
| `effects`   | `EffectsOptions`             | Visual effects                                        |
| `geometry`  | `string`                     | Preset geometry                                       |
| `rotation`  | `number`                     | Rotation in degrees                                   |
| `flipV`     | `boolean`                    | Flip vertically                                       |
| `flipH`     | `boolean`                    | Flip horizontally                                     |
| `animation` | `AnimationOptions`           | Animation settings                                    |

### Shape with Rich Text (textBody)

```json
{
  "x": "2cm",
  "y": "4cm",
  "width": "16cm",
  "height": "5cm",
  "textBody": {
    "children": [
      {
        "children": [{ "text": "Bold title", "bold": true, "size": 24, "color": "FFFFFF" }],
        "alignment": "center"
      },
      {
        "children": [{ "text": "Subtitle", "size": 16, "color": "CCCCCC" }]
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
{ "type": "none" }
```

## Outline

```json
{
  "color": "1565C0",
  "width": "2pt",
  "dashStyle": "solid",
  "capStyle": "flat",
  "compound": "single"
}
```

| Property    | Type                         | Values                                                        |
| ----------- | ---------------------------- | ------------------------------------------------------------- |
| `color`     | `string`                     | Hex color without `#`                                         |
| `width`     | `number \| UniversalMeasure` | Width in EMU or unit string                                   |
| `dashStyle` | `string`                     | `"solid"` \| `"dash"` \| `"dashDot"` \| `"dot"` \| `"lgDash"` |
| `capStyle`  | `string`                     | `"flat"` \| `"round"` \| `"square"`                           |
| `compound`  | `string`                     | `"single"` \| `"double"` \| `"thickThin"`                     |

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
  "bevelTop": { "w": 76200, "h": 76200 },
  "bevelBottom": { "w": 50800, "h": 50800 },
  "lighting": "threePt"
}
```

| Property      | Type             | Values                                                                           |
| ------------- | ---------------- | -------------------------------------------------------------------------------- |
| `extrusionH`  | `number`         | Extrusion height (EMU)                                                           |
| `material`    | `string`         | `"plastic"` \| `"metal"` \| `"matte"` \| `"warmMatte"` \| `"flat"` \| `"powder"` |
| `bevelTop`    | `{ w, h, prst }` | Bevel dimensions in EMU; optional preset                                         |
| `bevelBottom` | `{ w, h, prst }` | Bevel dimensions in EMU; optional preset                                         |
| `lighting`    | `string`         | `"flat"` \| `"threePt"` \| `"balanced"` \| `"soft"` \| `"harsh"`                 |

## Tables

```json
{
  "x": "2cm",
  "y": "8cm",
  "width": "24cm",
  "height": "10cm",
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
  "x": "4cm",
  "y": "4cm",
  "width": "12cm",
  "height": "8cm",
  "data": "<Uint8Array>",
  "imageType": "png"
}
```

## Charts

Charts use the pptx `ChartOptions` (which extends core's `ChartSpaceOptions`) and are embedded in slides as `{ chart: ChartOptions }`.

## Lines & Connectors

### Simple Line

```json
{
  "x1": "2cm",
  "y1": "5cm",
  "x2": "32cm",
  "y2": "5cm",
  "outline": { "color": "4472C4", "width": "2pt" }
}
```

### Connector (line with arrowheads)

```json
{
  "x1": "2cm",
  "y1": "5cm",
  "x2": "16cm",
  "y2": "5cm",
  "outline": {
    "color": "4472C4",
    "width": "2pt",
    "headEnd": { "type": "triangle" },
    "tailEnd": { "type": "triangle" }
  }
}
```

Arrowheads live inside `outline.headEnd` / `outline.tailEnd` (each a `LineEndOptions`):

| Property | Type     | Values                                                                          |
| -------- | -------- | ------------------------------------------------------------------------------- |
| `type`   | `string` | `"none"` \| `"triangle"` \| `"stealth"` \| `"diamond"` \| `"oval"` \| `"arrow"` |
| `width`  | `string` | `"sm"` \| `"med"` \| `"lg"`                                                     |
| `length` | `string` | `"sm"` \| `"med"` \| `"lg"`                                                     |

## Group Shapes

```json
{
  "x": "2cm",
  "y": "2cm",
  "width": "20cm",
  "height": "12cm",
  "rotation": 10,
  "children": [
    {
      "x": 0,
      "y": 0,
      "width": "9cm",
      "height": "5cm",
      "textBody": { "text": "A" },
      "fill": "4472C4"
    },
    {
      "x": "10cm",
      "y": 0,
      "width": "9cm",
      "height": "5cm",
      "textBody": { "text": "B" },
      "fill": "ED7D31"
    }
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
// parsePresentation accepts Uint8Array, ArrayBuffer, typed-array views, number[], base64 string
parsePresentation(readFileSync("input.pptx"))
// Returns PresentationOptions — pass to generatePresentation(opts) to modify and re-export
```

## Patching

### patchPresentation

Modify an existing `.pptx` template with ordered, source-preserving commands:

```ts
import { patchPresentation } from "@office-open/pptx";

const result = await patchPresentation({
  outputType: "nodebuffer",
  data: templateBuffer,
  commands: [
    {
      type: "replace",
      at: { field: "slides", route: [{ index: 0 }], text: { placeholder: "title" } },
      value: "Updated",
    },
    {
      type: "update",
      at: { field: "slides", route: [{ index: 0 }] },
      set: { name: "Overview" },
    },
    {
      type: "insert",
      at: { field: "slides" },
      position: "end",
      value: [
        {
          children: [
            { shape: { x: 0, y: 0, width: 200, height: 100, textBody: { text: "Appended" } } },
          ],
        },
      ],
    },
    { type: "move", at: { field: "slides", route: [{ index: 1 }] }, to: "start" },
    { type: "remove", at: { field: "slides", route: [{ index: 1 }] } },
    {
      type: "insert",
      at: { field: "slides", route: [{ index: 0 }, { collection: "comments" }] },
      value: [{ author: "Alice", text: "Check this shape.", x: 10, y: 20 }],
    },
  ],
  placeholderDelimiters: { start: "{{", end: "}}" },
});
```

Each command uses `at: { field, route, text }`; slide routes accept `index`, `id`, or `name`, and optional `text` selects placeholder or literal text. Failed patches throw a structured `PatchError` with a diagnostic containing `part`, `path`, `operation`, `reason`, and `message`.

| Option                  | Type                                                     | Default      | Description                     |
| ----------------------- | -------------------------------------------------------- | ------------ | ------------------------------- |
| `outputType`            | `string`                                                 | —            | Output format                   |
| `data`                  | `Buffer \| Uint8Array \| ...`                            | —            | Input .pptx file                |
| `commands`              | `readonly PresentationPatchCommand[]`                    | —            | Ordered source-preserving edits |
| `packageVariant`        | `"standard" \| "macro" \| "template" \| "macroTemplate"` | `"standard"` | Output package variant          |
| `placeholderDelimiters` | `{ start, end }`                                         | `{ {{, }} }` | Custom delimiters               |

| Command   | Value field | Purpose                                                    |
| --------- | ----------- | ---------------------------------------------------------- |
| `insert`  | `value`     | Add canonical items to an array field or nested collection |
| `replace` | `value`     | Replace a canonical field/node or selected text            |
| `update`  | `set`       | Deep-merge object fields; arrays are replaced as a whole   |
| `remove`  | —           | Remove a canonical node or selected text                   |
| `move`    | `to`        | Reorder an item in a canonical array                       |
