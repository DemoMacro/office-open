# Measurement Units & Converters

OOXML uses several measurement systems. This reference covers all units and conversion utilities.

## Unit Reference

| Unit        | Used In                         | Conversion          |
| ----------- | ------------------------------- | ------------------- |
| Twips       | Page margins, paragraph spacing | 1 inch = 1440 twips |
| EMU         | Images, shapes, offsets         | 1 inch = 914400 EMU |
| Half-points | Font sizes                      | Size 24 = 12pt      |
| 60000ths°   | Angles, rotation, effects       | 45° = 2700000       |
| 1000ths %   | Alpha/opacity in effects        | 50% = 50000         |

## Common Values

### Page Sizes (Twips)

| Size   | Width | Height |
| ------ | ----- | ------ |
| A4     | 11906 | 16838  |
| Letter | 12240 | 15840  |
| A3     | 16838 | 23811  |
| B5     | 10318 | 14570  |

### Margins (Twips)

| Margin   | Default |
| -------- | ------- |
| 1 inch   | 1440    |
| 0.5 inch | 720     |
| 2.54 cm  | 1440    |

### Font Sizes (Half-points)

| Size | Half-points |
| ---- | ----------- |
| 10pt | 20          |
| 11pt | 22          |
| 12pt | 24          |
| 14pt | 28          |
| 16pt | 32          |
| 24pt | 48          |
| 36pt | 72          |

### Angles (60000ths of a degree)

| Degrees | Value    |
| ------- | -------- |
| 0°      | 0        |
| 45°     | 2700000  |
| 90°     | 5400000  |
| 180°    | 10800000 |
| 270°    | 16200000 |
| 360°    | 21600000 |

## Converter Functions

Available from `@office-open/core`:

```ts
import {
  convertMillimetersToTwip,
  convertInchesToTwip,
  convertPixelsToEmu,
  convertEmuToPixels,
  convertTwipToEMU,
} from "@office-open/core";

// mm → twips
const margin = convertMillimetersToTwip(25.4); // 1440

// inches → twips
const indent = convertInchesToTwip(1); // 1440

// pixels → EMU
const emu = convertPixelsToEmu(200); // Used for shape/image dimensions
```

## Pixel to EMU Conversion

Shapes and images in PPTX use pixel values in the API, which are internally converted to EMU:

- 1 pixel = 9525 EMU (at 96 DPI)
- The `x`, `y`, `width`, `height` properties accept pixel values
- Internal conversion: `convertPixelsToEmu(pixels)`
