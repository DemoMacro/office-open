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

## Bundle Size

|      | @office-open/xml |
| ---- | ---------------: |
| gzip |      **4.22 kB** |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
