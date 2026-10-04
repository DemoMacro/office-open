# @office-open/odt

![npm version](https://img.shields.io/npm/v/@office-open/odt)
![npm downloads](https://img.shields.io/npm/dw/@office-open/odt)
![npm license](https://img.shields.io/npm/l/@office-open/odt)

> Generate and parse OpenDocument Text `.odt` files using canonical `DocumentOptions` JSON.

## Installation

```bash
pnpm add @office-open/odt
```

## Quick Start

```typescript
import { generateOdt, parseOdt } from "@office-open/odt";

const odt = generateOdt({
  sections: [{ children: [{ paragraph: "Hello ODT" }] }],
});
const options = parseOdt(odt);
```

## API

- `generateOdt(options, context?)` — generate an `.odt` package from `DocumentOptions`
- `parseOdt(data, options?)` — parse an `.odt` package into `DocumentOptions`
- `OdtParseError` — structured parse failure with part, path, and reason diagnostics

## Documentation

- [Documentation](https://www.office-open.com/en/getting-started/more-formats) — additional format guides and examples
- [Report Issues](https://github.com/DemoMacro/office-open/issues) — bug reports and feature requests

## License

- [MIT](https://github.com/DemoMacro/office-open/blob/main/LICENSE) © [Demo Macro](https://www.demomacro.com/)
