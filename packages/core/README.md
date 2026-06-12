# @office-open/core

![npm version](https://img.shields.io/npm/v/@office-open/core)
![npm downloads](https://img.shields.io/npm/dw/@office-open/core)
![npm license](https://img.shields.io/npm/l/@office-open/core)

> Shared OOXML infrastructure — descriptors, DrawingML, charts, SmartArt, OPC, validators, and converters.

## Features

- **Descriptor Runtime** — `CustomDescriptor<T>` and `ElementDescriptor<T>` for bidirectional XML stringify/parse
- **DrawingML** — Fills, outlines, effects, geometry, text body, scene 3D, and more
- **Chart Components** — Shared chart types (bar, line, pie, area, scatter) and chart collection
- **SmartArt Components** — Data model, tree-to-model converter, layout/style/color definitions
- **OPC (Open Packaging)** — Content types, relationships, and ZIP packaging
- **Value Validators** — Runtime validation for OOXML spec types (ST_HexColor, ST_OnOff, ST_DecimalNumber, etc.)
- **Unit Converters** — TWIP and EMU conversions (mm/in/pt/px)
- **ID Generators** — Sequential numeric IDs, random IDs, SHA-1 hash, UUID v4
- **Template Patching** — Placeholder replacement, XML traverser, and token replacer
- **OOXML Compliance** — All types verified against ISO/IEC 29500-4 XSD schemas

## Installation

```bash
# pnpm
pnpm add @office-open/core

# npm
npm install @office-open/core

# yarn
yarn add @office-open/core

# bun
bun add @office-open/core
```

## Quick Start

```typescript
import {
  hexColorValue,
  decimalNumber,
  convertMillimetersToTwip,
  convertPixelsToEmu,
  convertInchesToEmu,
  convertPointsToEmu,
  uniqueNumericIdCreator,
} from "@office-open/core";

// Value validators
hexColorValue("#FF0000"); // → "FF0000"
hexColorValue("auto"); // → "auto"
decimalNumber(10.7); // → 10

// Unit converters
convertMillimetersToTwip(25.4); // → 1440 (1 inch)
convertPixelsToEmu(100); // → 952500
convertInchesToEmu(1); // → 914400
convertPointsToEmu(12); // → 152400

// ID generators
const gen = uniqueNumericIdCreator();
gen(); // → 1, 2, 3, ...
```

## OOXML Schema Compliance

| Type               | XSD Reference                                  |
| ------------------ | ---------------------------------------------- |
| `ThemeColor`       | `ST_ThemeColor` (17 values)                    |
| `ThemeFont`        | `ST_Theme` (8 values)                          |
| `UniversalMeasure` | `ST_UniversalMeasure` (mm, cm, in, pt, pc, pi) |
| `Percentage`       | `ST_Percentage`                                |
| `hexColorValue`    | `ST_HexColor` (auto + 3-byte hexBinary)        |
| `decimalNumber`    | `ST_DecimalNumber`                             |

## Exports

| Path                          | Contents                                                                  |
| ----------------------------- | ------------------------------------------------------------------------- |
| `@office-open/core`           | Descriptors, validators, converters, ID generators, OPC, password hashing |
| `@office-open/core/values`    | Validators + ThemeColor/ThemeFont only                                    |
| `@office-open/core/chart`     | Chart types, series data, chart collection, title                         |
| `@office-open/core/smartart`  | SmartArt data model, tree-to-model, definitions                           |
| `@office-open/core/drawingml` | DrawingML fills, outlines, effects, geometry, text body                   |
| `@office-open/core/patch`     | Template patching utilities (replacer, traverser, token replacer)         |
| `@office-open/core/theme`     | Theme definitions and color schemes                                       |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
