# @office-open/core

![npm version](https://img.shields.io/npm/v/@office-open/core)
![npm downloads](https://img.shields.io/npm/dw/@office-open/core)
![npm license](https://img.shields.io/npm/l/@office-open/core)

> Shared OOXML infrastructure: XmlComponent, value validators, unit converters.

## Features

- **XmlComponent Framework** - Base classes for building OOXML element trees with dynamic namespace support
- **Value Validators** - Runtime validation for OOXML spec types (ST_HexColor, ST_OnOff, ST_DecimalNumber, etc.)
- **Unit Converters** - Millimeter/inch to TWIP conversion
- **ID Generators** - Sequential numeric IDs, nanoid, SHA-1 hash, UUID v4
- **OOXML Compliance** - All types verified against ISO/IEC 29500-4 XSD schemas

## Installation

```bash
# Install with npm
$ npm install @office-open/core

# Install with pnpm
$ pnpm add @office-open/core
```

## Quick Start

```typescript
import {
    OnOffElement,
    StringValueElement,
    BuilderElement,
    StringContainer,
    hexColorValue,
    decimalNumber,
    convertMillimetersToTwip,
    uniqueNumericIdCreator,
} from "@office-open/core";

// CT_OnOff — dynamic namespace prefix per XSD spec
new OnOffElement("w:b", false);
// → { "w:b": { _attr: { "w:val": false } } }

new OnOffElement("m:hideBot", false);
// → { "m:hideBot": { _attr: { "m:val": false } } }

// Builder with attributes + children
new BuilderElement({
    name: "w:r",
    attributes: { lang: { key: "xml:lang", value: "en-US" } },
    children: [new StringContainer("w:t", "Hello")],
});

// Value validators
hexColorValue("#FF0000"); // → "FF0000"
hexColorValue("auto"); // → "auto"
decimalNumber(10.7); // → 10

// Unit converters
convertMillimetersToTwip(25.4); // → 1440 (1 inch)

// ID generators
const gen = uniqueNumericIdCreator();
gen(); // → 1, 2, 3, ...
```

## OOXML Schema Compliance

| Type                 | XSD Reference                                  |
| -------------------- | ---------------------------------------------- |
| `ThemeColor`         | `ST_ThemeColor` (17 values)                    |
| `ThemeFont`          | `ST_Theme` (8 values)                          |
| `UniversalMeasure`   | `ST_UniversalMeasure` (mm, cm, in, pt, pc, pi) |
| `Percentage`         | `ST_Percentage`                                |
| `hexColorValue`      | `ST_HexColor` (auto + 3-byte hexBinary)        |
| `OnOffElement`       | `CT_OnOff` (dynamic namespace prefix)          |
| `HpsMeasureElement`  | `CT_HpsMeasure`                                |
| `StringValueElement` | `CT_String`                                    |

## Exports

| Path                       | Contents                                            |
| -------------------------- | --------------------------------------------------- |
| `@office-open/core`        | XmlComponent, validators, converters, ID generators |
| `@office-open/core/values` | Validators + ThemeColor/ThemeFont only              |

## Benchmark

| Operation                     | hz    |
| ----------------------------- | ----- |
| `decimalNumber`               | ~21M  |
| `hexColorValue` (6-char hex)  | ~13M  |
| `uniqueNumericIdCreator`      | ~21M  |
| `uniqueId` (nanoid)           | ~3.3M |
| `uniqueUuid`                  | ~2M   |
| `OnOffElement (true)`         | ~13M  |
| `OnOffElement (false)`        | ~1.7M |
| `BuilderElement` (attributes) | ~2M   |
| `BuilderElement` (children)   | ~2.3M |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
