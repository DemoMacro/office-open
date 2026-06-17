You are a senior TypeScript developer.

## Project

**office-open** is a monorepo for generating and parsing Office Open XML documents (.docx, .pptx, .xlsx) with JS/TS. Declarative API, works in Node.js and browsers. Bidirectional: stringify (JSON → XML) and parse (XML → JSON).

## Monorepo Structure

```
packages/
  core/   — @office-open/core (descriptor runtime, DrawingML factories, chart/smartart, OPC)
  xml/    — @office-open/xml (XML parsing/serialization)
  docx/   — @office-open/docx (DOCX generation & parsing)
  pptx/   — @office-open/pptx (PPTX generation & parsing)
  xlsx/   — @office-open/xlsx (XLSX generation & parsing)
ooxml-schemas/  — OOXML XSD schemas (golden source of truth)
```

## Package Layout (all format packages follow this pattern)

```
packages/<format>/src/
  index.ts        — Public API
  compiler.ts     — compileDocument/Presentation/Workbook → Zippable
  context.ts      — XxxWriteContext + XxxReadContext
  generate.ts     — generateDocument/Presentation/Workbook() entry
  parse.ts        — parseDocument/Presentation/Workbook() entry
  patch.ts        — patchDocument/Presentation/Workbook() entry
                    (single file when simple; use patch/ directory if complex)
  parts/          — One module per OOXML XML part
                    Simple parts = single file (e.g. settings.ts)
                    Complex parts = directory (e.g. document/)
  shared/         — Cross-part shared types (used by 2+ parts)
  util/           — Helpers
```

### parts/ — OOXML Part Modules

Each part module exports its types and descriptor together:

```typescript
// parts/settings.ts — simple part (single file)
export interface SettingsOptions { ... }
export const settingsDesc: CustomDescriptor<SettingsOptions> = {
  kind: "custom",
  stringify(opts, ctx) { ... },
  parse(el, ctx) { ... },
};
```

Complex parts use a directory with sub-modules (e.g. `parts/document/` contains body.ts, run.ts, paragraph.ts, math.ts, etc.).

### shared/ — Cross-Part Types

Types used by 2+ parts: RunOptions, ParagraphOptions, BorderOptions, Media, etc. Follow the colocation rule: single-use types live in their part module; multi-use types go to shared/.

## Build & Test

- **Build**: `pnpm build` (all) or `cd packages/<pkg> && pnpm build` (one)
- **Test**: `vp test run` (per package)
- **Lint**: `pnpm check` (root, runs vp check across all)
- **Validate**: `pnpm tsx scripts/validate.ts` (XSD validation)

## Descriptor System

Core infrastructure at `packages/core/src/descriptor/`:

- `CustomDescriptor<T>` — `{ kind: "custom", stringify(opts, ctx): string, parse(el, ctx): T }`
- Runtime: `stringify(desc, value, ctx)` and `parse(desc, element, ctx)`

Naming: `<part>Desc` for descriptors (e.g. `settingsDesc`, `slideDesc`).

## OOXML Schemas

`ooxml-schemas/transitional/` is the primary reference. Key files:

- `wml.xsd` — WordprocessingML, `pml.xsd` — PresentationML
- `sml.xsd` — SpreadsheetML, `dml-main.xsd` — DrawingML

## Code Conventions

Full standards in [CONTRIBUTING.md](./CONTRIBUTING.md). Quick reference:

- **Naming**: `<part>Desc` descriptors, `<Part>Options` interfaces, `stringify*()` / `parse*()` / `patch*()` helpers
- **Properties**: full English words (camelCase); reference elements → `*Reference`; OOXML attribute tokens (`id`/`idx`/`numFmt`/`fontId`/…) preserved verbatim; never compound abbreviations like `lnIdx` → `lineReferenceIndex` — see [CONTRIBUTING.md](./CONTRIBUTING.md#property-naming)
- **Constants**: `as const` objects, SCREAMING_SNAKE_CASE keys, lowercase values
- **Files**: kebab-case, no `I` prefix on interfaces, no `readonly` on Options properties
- **Loops**: `for...of` default, `.map()` only when returning new array
- **XML generation**: string concatenation via template literals, no intermediate object trees

## Behavioral Guidelines

- State assumptions explicitly. If uncertain, ask before implementing.
- No features beyond what was asked. No speculative abstractions.
- Touch only what you must. Match existing style.
- Transform tasks into verifiable goals. Loop until verified.
