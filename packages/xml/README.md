# @office-open/xml

![npm version](https://img.shields.io/npm/v/@office-open/xml)
![npm downloads](https://img.shields.io/npm/dw/@office-open/xml)
![npm license](https://img.shields.io/npm/l/@office-open/xml)
![zero dependencies](https://img.shields.io/badge/dependencies-0-green)

> XML parsing and serialization for Office Open XML. Zero dependencies, pure TypeScript.

## Features

- **Zero Dependencies** - No external runtime dependencies, pure TypeScript implementation
- **parse() / stringify()** - XML string ↔ Element tree, OOXML-optimized
- **Element Type** - Tolerant element model for round-tripping Office Open XML parts
- **escapeXml() / unescapeXml()** - Low-level XML entity escaping
- **OOXML Optimized** - Implements the options needed for Office Open XML document generation and parsing

## Installation

```bash
# pnpm
pnpm add @office-open/xml

# npm
npm install @office-open/xml

# yarn
yarn add @office-open/xml

# bun
bun add @office-open/xml
```

## Quick Start

```typescript
import { parse, stringify } from "@office-open/xml";

// Parse XML to an Element tree
const doc = parse("<w:t>Hello</w:t>");

// Serialize an Element tree back to XML
const xml = stringify(doc);
```

## API

### parse(xmlString, options?)

Parse an XML string into an `Element` tree. Options include `compact`, `trim`, `nativeType`, `captureSpacesBetweenElements`, the `ignore*` flags, and the `*Fn` transformation hooks.

### stringify(element, options?)

Serialize an `Element` tree to an XML string. Options include `spaces` (indentation), the `ignore*` flags, and the `*Fn` hooks.

### escapeXml(str) / unescapeXml(str)

Low-level XML entity escaping and unescaping.

### Element

The tolerant element type used across all office-open packages:

```typescript
interface Element {
  declaration?: { attributes?: DeclarationAttributes };
  attributes?: Attributes;
  type?: string;
  name?: string;
  text?: string | number | boolean;
  cdata?: string;
  comment?: string;
  elements?: Element[];
}
```

## Benchmark

Performance vs [txml](https://github.com/TobiasNickel/tXml), [xml-js](https://github.com/nashwaan/xml-js), and [xml](https://github.com/dylang/node-xml) (higher ops/s is better, Windows 11 / Node 24; run with `pnpm exec vp test bench src/xml.bench.ts` in this package). `@office-open/xml` is a drop-in replacement for xml-js and xml. The `xml` (npm) package is generation-only (no parser), so it only appears under stringify.

txml skips entity encoding by default (`encodeEntities: false`), which emits invalid XML when text contains `&`, `<`, or `>`; the `txml` column shows that raw mode and the `txml (entities)` column is the output-equivalent mode.

**parse() — XML string → Element tree**

| Scenario      | @office-open/xml |            txml |       xml-js |
| ------------- | ---------------: | --------------: | -----------: |
| simple XML    |  1,420,676 ops/s | 1,077,871 ops/s | 99,049 ops/s |
| complex OOXML |    434,118 ops/s |   393,714 ops/s | 52,849 ops/s |

**stringify() — Element tree → XML string**

| Scenario       | @office-open/xml | txml (entities) |            txml |        xml-js |     xml (npm) |
| -------------- | ---------------: | --------------: | --------------: | ------------: | ------------: |
| simple element |  2,039,248 ops/s | 1,339,180 ops/s | 3,188,605 ops/s | 207,318 ops/s | 299,226 ops/s |
| complex OOXML  |    579,213 ops/s |   465,948 ops/s | 1,409,039 ops/s | 137,595 ops/s | 186,069 ops/s |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
