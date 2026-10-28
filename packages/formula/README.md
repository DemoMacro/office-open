# @office-open/formula

![npm version](https://img.shields.io/npm/v/@office-open/formula)
![npm downloads](https://img.shields.io/npm/dw/@office-open/formula)
![npm license](https://img.shields.io/npm/l/@office-open/formula)

> Parse, analyze, rewrite, and evaluate spreadsheet formulas in TypeScript and JavaScript — no Microsoft Office required. Built for AI agents and hand-written code alike; runs in Node.js, browsers, Deno, and Bun.

## Features

- 🧮 **Formula Parsing** — Excel and OpenFormula references, structured tables, 3D ranges, unions, intersections, spill references, and arrays
- 📚 **Function Catalog** — 667 researched functions and compatibility aliases across Excel, OpenFormula, LibreOffice, OnlyOffice, and WPS
- ⚡ **Evaluation** — scalar and matrix values, lazy branches, lambda scopes, date systems, and structured Excel errors
- 🔗 **Workbook Integration** — dense or sparse cells, defined names, tables, hidden rows, dynamic arrays, pivot data, web services, macros, SQL, cubes, DDE, and RTD host capabilities
- 🧭 **Analysis Tools** — dependency collection, graph construction, cycle detection, topological evaluation, shared-formula resolution, and relative-reference rewriting
- 🧩 **Typed Host Integration** — structured workbook snapshots, dynamic arrays, lookup providers, cube and pivot requests, web services, macros, SQL, DDE, and RTD

## Installation

```bash
# pnpm
pnpm add @office-open/formula

# npm
npm install @office-open/formula

# yarn
yarn add @office-open/formula

# bun
bun add @office-open/formula
```

## Quick Start

```typescript
import {
  buildDependencyGraph,
  collectDependencies,
  evaluateFormula,
  parseFormula,
  topologicalFormulaOrder,
} from "@office-open/formula";

const workbook = {
  getSheetId: (name) => (name === "Sheet1" ? 0 : undefined),
  getSheetName: (id) => (id === 0 ? "Sheet1" : undefined),
  getCell: (sheet, row, column) =>
    sheet === 0 && row === 1 && column === 1 ? { value: 2, formula: undefined } : undefined,
  getDefinedFormula: () => undefined,
  getTableName: () => undefined,
  getRowCount: () => 1_048_576,
  getColumnCount: () => 16_384,
};

const evaluated = evaluateFormula("=A1*2", workbook, { sheet: 0, row: 1, column: 2 });
const ast = parseFormula("=A1*2");
const dependencies = collectDependencies(ast);
const graph = buildDependencyGraph([
  { id: "Sheet1!B1", ast, cell: { sheet: "Sheet1", row: 1, column: 2 } },
]);

console.log(evaluated.value, dependencies, [...topologicalFormulaOrder(graph)]);
```

## API

- `parseFormula(source, options?)` — parse Excel or OpenFormula syntax into a typed AST
- `stringifyFormula(ast, options?)` — serialize an AST back to formula text
- `evaluateFormula(formula, workbook, options?)` — evaluate a formula against a typed workbook
- `evaluateFormulaAst(ast, workbook, options?)` — evaluate a parsed AST without reparsing
- `collectDependencies(ast)` — collect cell, range, name, table, and unknown dependencies
- `buildDependencyGraph(entries, options?)` — build a dependency graph for workbook evaluation
- `detectFormulaCycles(graph)` / `topologicalFormulaOrder(graph)` — validate and order evaluation
- `shiftFormula(formula, options?)` / `shiftFormulaAst(ast, options?)` — rewrite relative references
- `getFunction(name)` / `getFunctionMetadata(name)` / `listFunctionNames()` — inspect the function registry
- `FormulaError` / `serializeFormulaError(value)` — distinguish runtime errors from error-like text

## Errors and Limitations

- Malformed syntax throws a structured `FormulaParseError` with source position and dialect information.
- Runtime evaluation errors serialize to canonical spreadsheet tokens by default; pass `{ errors: "object" }` to receive `FormulaError` instances.
- Host-dependent functions return canonical errors when the workbook does not provide their typed capability.
- Unknown host and numeric-library exceptions propagate rather than being masked as `#VALUE!`.
- Dynamic arrays allocate at most 2,000,000 cells; larger dimensions return `#VALUE!`. Supply `getCells` for sparse ranges so operations such as `COUNT` and `COUNTA` can scan actual entries without materializing the full rectangle.

## Documentation

- [Documentation](https://www.office-open.com/en/) — guides, API reference, and examples
- [Report Issues](https://github.com/DemoMacro/office-open/issues) — bug reports and feature requests

## Related Packages

- [office-open](https://www.npmjs.com/package/office-open) — all formats + CLI + AI SDK tools in one install
- [@office-open/xlsx](https://www.npmjs.com/package/@office-open/xlsx) — Excel (.xlsx)
- [@office-open/core](https://www.npmjs.com/package/@office-open/core) — shared OOXML infrastructure

## License

- [MIT](https://github.com/DemoMacro/office-open/blob/main/LICENSE) &copy; [Demo Macro](https://www.demomacro.com/)
