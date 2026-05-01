# @office-open/xml

![npm version](https://img.shields.io/npm/v/@office-open/xml)
![npm downloads](https://img.shields.io/npm/dw/@office-open/xml)
![npm license](https://img.shields.io/npm/l/@office-open/xml)
![zero dependencies](https://img.shields.io/badge/dependencies-0-green)

> XML parsing and serialization for Office Open XML. Zero dependencies, drop-in replacement for `xml` + `xml-js`.

## Features

- **Zero Dependencies** - No external runtime dependencies, pure TypeScript implementation
- **xml() Serialization** - Drop-in replacement for the `xml` package
- **xml2js() Parsing** - Drop-in replacement for `xml-js` XML parsing
- **js2xml() Stringifying** - Drop-in replacement for `xml-js` JS-to-XML conversion
- **toElement() Direct Convert** - Direct conversion from xml object format to xml-js Element, 16-33x faster than the xml→xml2js bridge
- **Complete Type Definitions** - Full type compatibility with `xml` and `xml-js`, import without changes
- **OOXML Optimized** - Implements all options needed for Office Open XML document generation

## Installation

```bash
# Install with npm
$ npm install @office-open/xml

# Install with pnpm
$ pnpm add @office-open/xml
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
| Simple element          |     4,103,029 hz | 734,685 hz | **5.58x** |
| Nested element          |       940,333 hz | 321,791 hz | **2.92x** |
| Nested with declaration |       800,517 hz | 262,125 hz | **3.05x** |

### Parsing (xml2js)

| Scenario           | @office-open/xml |    xml-js |    Speedup |
| ------------------ | ---------------: | --------: | ---------: |
| Simple XML         |     1,002,659 hz | 83,057 hz | **12.07x** |
| Complex OOXML      |       357,081 hz | 48,687 hz |  **7.34x** |
| With captureSpaces |       361,555 hz | 44,868 hz |  **8.06x** |

### Stringifying (js2xml)

| Scenario       | @office-open/xml |     xml-js |   Speedup |
| -------------- | ---------------: | ---------: | --------: |
| Simple element |       780,070 hz | 161,521 hz | **4.83x** |
| Complex OOXML  |       276,878 hz | 106,527 hz | **2.60x** |

### Direct Conversion (toElement vs bridge)

| Scenario |   toElement() | xml() + xml2js() bridge |    Speedup |
| -------- | ------------: | ----------------------: | ---------: |
| Simple   | 14,119,071 hz |              853,290 hz | **16.55x** |
| Nested   |  3,745,934 hz |              422,217 hz |  **8.87x** |

## Bundle Size

|      | @office-open/xml | xml + xml-js |
| ---- | ---------------: | -----------: |
| gzip |      **4.22 kB** |       ~15 kB |

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
