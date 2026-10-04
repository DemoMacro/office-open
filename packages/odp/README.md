# @office-open/odp

![npm version](https://img.shields.io/npm/v/@office-open/odp)
![npm downloads](https://img.shields.io/npm/dw/@office-open/odp)
![npm license](https://img.shields.io/npm/l/@office-open/odp)

> Generate and parse OpenDocument Presentation `.odp` files using canonical `PresentationOptions` JSON.

## Installation

```bash
pnpm add @office-open/odp
```

## Quick Start

```typescript
import { generateOdp, parseOdp } from "@office-open/odp";

const odp = generateOdp({
  slides: [{ children: [] }],
});
const options = parseOdp(odp);
```

## API

- `generateOdp(options, context?)` — generate an `.odp` package from `PresentationOptions`
- `parseOdp(data, options?)` — parse an `.odp` package into `PresentationOptions`
- `OdpParseError` — structured parse failure with part, path, and reason diagnostics

## Documentation

- [Documentation](https://www.office-open.com/en/getting-started/more-formats) — additional format guides and examples
- [Report Issues](https://github.com/DemoMacro/office-open/issues) — bug reports and feature requests

## License

- [MIT](https://github.com/DemoMacro/office-open/blob/main/LICENSE) © [Demo Macro](https://www.demomacro.com/)
