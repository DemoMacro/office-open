# @office-open/ods

![npm version](https://img.shields.io/npm/v/@office-open/ods)
![npm downloads](https://img.shields.io/npm/dw/@office-open/ods)
![npm license](https://img.shields.io/npm/l/@office-open/ods)

> Generate and parse OpenDocument Spreadsheet `.ods` files using canonical `WorkbookOptions` JSON.

## Installation

```bash
pnpm add @office-open/ods
```

## Quick Start

```typescript
import { generateOds, parseOds } from "@office-open/ods";

const ods = generateOds({
  worksheets: [{ name: "Sheet 1", rows: [{ cells: [{ value: 42 }] }] }],
});
const options = parseOds(ods);
```

## API

- `generateOds(options, context?)` — generate an `.ods` package from `WorkbookOptions`
- `parseOds(data, options?)` — parse an `.ods` package into `WorkbookOptions`
- `OdsParseError` — structured parse failure with part, path, and reason diagnostics

## Documentation

- [Documentation](https://www.office-open.com/en/getting-started/more-formats) — additional format guides and examples
- [Report Issues](https://github.com/DemoMacro/office-open/issues) — bug reports and feature requests

## License

- [MIT](https://github.com/DemoMacro/office-open/blob/main/LICENSE) © [Demo Macro](https://www.demomacro.com/)
