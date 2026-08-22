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

Performance vs [txml](https://github.com/TobiasNickel/tXml), [xml-js](https://github.com/nashwaan/xml-js), and [xml](https://github.com/dylang/node-xml) (higher ops/s is better, Windows 11; each scenario runs under both Node 24 and Bun 1.4). `@office-open/xml` is a drop-in replacement for xml-js and xml. The `xml` (npm) package is generation-only (no parser), so it only appears under stringify. `Bun.XML` is Bun-only and its compact parse mode is lossy (`#text` runs concatenate, same-name children collapse into keyed arrays), so it is a throughput reference rather than a drop-in option.

txml skips entity encoding by default (`encodeEntities: false`), which emits invalid XML when text contains `&`, `<`, or `>`; the `txml` column shows that raw mode and the `txml (entities)` column is the output-equivalent mode.

**parse() — XML string → Element tree**

| Scenario      | Runtime | @office-open/xml | txml          | xml-js        | Bun.XML.parse |
| ------------- | ------- | ---------------- | ------------- | ------------- | ------------- |
| simple XML    | Node 24 | 1,301,742 ops/s  | 983,046 ops/s | 93,951 ops/s  | —             |
|               | Bun 1.4 | 1,906,525 ops/s  | 616,562 ops/s | 161,894 ops/s | 828,554 ops/s |
| complex OOXML | Node 24 | 423,536 ops/s    | 388,743 ops/s | 49,771 ops/s  | —             |
|               | Bun 1.4 | 483,512 ops/s    | 241,036 ops/s | 63,029 ops/s  | 451,177 ops/s |

**stringify() — Element tree → XML string**

| Scenario       | Runtime | @office-open/xml | txml (entities) | txml            | xml-js        | xml (npm)     | Bun.XML.stringify |
| -------------- | ------- | ---------------- | --------------- | --------------- | ------------- | ------------- | ----------------- |
| simple element | Node 24 | 2,049,617 ops/s  | 1,255,456 ops/s | 2,765,770 ops/s | 190,783 ops/s | 303,724 ops/s | —                 |
|                | Bun 1.4 | 2,588,728 ops/s  | 1,721,498 ops/s | 4,248,377 ops/s | 438,430 ops/s | 483,526 ops/s | 584,746 ops/s     |
| complex OOXML  | Node 24 | 583,811 ops/s    | 447,274 ops/s   | 1,326,397 ops/s | 130,335 ops/s | 172,258 ops/s | —                 |
|                | Bun 1.4 | 510,448 ops/s    | 528,375 ops/s   | 2,303,173 ops/s | 226,645 ops/s | 292,534 ops/s | 251,444 ops/s     |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
