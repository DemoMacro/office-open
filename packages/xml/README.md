# @office-open/xml

![npm version](https://img.shields.io/npm/v/@office-open/xml)
![npm downloads](https://img.shields.io/npm/dw/@office-open/xml)
![npm license](https://img.shields.io/npm/l/@office-open/xml)
![zero dependencies](https://img.shields.io/badge/dependencies-0-green)

> XML parsing and serialization for Office Open XML. Zero dependencies, drop-in replacement for xml + xml-js.

## Features

- **Zero Dependencies** - No external runtime dependencies, pure TypeScript implementation
- **xml() Serialization** - Drop-in replacement for the `xml` package
- **xml2js() Parsing** - Drop-in replacement for `xml-js` XML parsing
- **js2xml() Stringifying** - Drop-in replacement for `xml-js` JS-to-XML conversion
- **toElement() Direct Convert** - Direct conversion from xml object format to xml-js Element, 10-19x faster than the xml→xml2js bridge
- **Complete Type Definitions** - Full type compatibility with `xml` and `xml-js`, import without changes
- **OOXML Optimized** - Implements all options needed for Office Open XML document generation

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

## Migration from xml + xml-js

Replace your existing imports:

```typescript
// Before
import xml from "xml";
import { xml2js, js2xml } from "xml-js";
import type { Element } from "xml-js";

// After
import { xml, xml2js, js2xml } from "@office-open/xml";
import type { Element } from "@office-open/xml";
```

No other code changes needed. All options and output formats are compatible.

## Quick Start

```typescript
import { xml, xml2js, js2xml, toElement } from "@office-open/xml";

// Serialize JS objects to XML
const xmlStr = xml({ "w:p": [{ _attr: { "w:val": "1" } }, { "w:r": [{ "w:t": "Hello" }] }] });
// <w:p w:val="1"><w:r><w:t>Hello</w:t></w:r></w:p>

// Parse XML to JS objects
const parsed = xml2js("<w:t>Hello</w:t>", { compact: false });

// Convert JS objects back to XML
const output = js2xml(parsed);

// Direct conversion (faster than xml → xml2js bridge)
const element = toElement({
  "w:p": [{ _attr: { "w:val": "1" } }, { "w:r": [{ "w:t": "Hello" }] }],
});
```

## API

### xml(input, options?)

Serialize JavaScript objects to XML string. Compatible with the `xml` package.

### xml2js(xmlString, options?)

Parse XML string to JavaScript object. Compatible with `xml-js`.

### js2xml(jsObject, options?)

Convert JavaScript object (xml-js Element format) to XML string. Compatible with `xml-js`.

### json2xml(jsObject, options?)

Alias for `js2xml`.

### xml2json(xmlString, options?)

Convenience function that returns `JSON.stringify(xml2js(xmlString, options))`.

### toElement(input)

Direct conversion from xml object format to xml-js Element format. Much faster than the `xml() → xml2js()` bridge path.

### escapeXml(str) / escapeAttributeValue(str)

Low-level XML entity escaping functions.

## Benchmark

Performance comparison against original `xml` (1.0.1) and `xml-js` (1.6.11) packages:

### Serialization (xml)

| Scenario                | @office-open/xml |        xml |   Speedup |
| ----------------------- | ---------------: | ---------: | --------: |
| Simple element          |     5,440,771 hz | 805,545 hz | **6.75x** |
| Nested element          |     1,050,272 hz | 315,184 hz | **3.33x** |
| Nested with declaration |       967,945 hz | 275,684 hz | **3.51x** |

### Parsing (xml2js)

| Scenario           | @office-open/xml |     xml-js |   Speedup |
| ------------------ | ---------------: | ---------: | --------: |
| Simple XML         |       869,965 hz | 100,507 hz | **8.66x** |
| Complex OOXML      |       346,440 hz |  53,621 hz | **6.46x** |
| With captureSpaces |       344,586 hz |  52,414 hz | **6.57x** |

### Stringifying (js2xml)

| Scenario       | @office-open/xml |     xml-js |   Speedup |
| -------------- | ---------------: | ---------: | --------: |
| Simple element |       793,710 hz | 207,730 hz | **3.82x** |
| Complex OOXML  |       366,515 hz | 135,815 hz | **2.70x** |

### Direct Conversion (toElement vs bridge)

| Scenario |   toElement() | xml() + xml2js() bridge |    Speedup |
| -------- | ------------: | ----------------------: | ---------: |
| Simple   | 15,471,913 hz |            1,073,016 hz | **14.42x** |
| Nested   |  4,278,499 hz |              441,468 hz |  **9.69x** |

## Bundle Size

|      | @office-open/xml | xml + xml-js |
| ---- | ---------------: | -----------: |
| gzip |      **4.22 kB** |       ~15 kB |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
