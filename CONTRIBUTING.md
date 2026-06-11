# Contributing to office-open

Thank you for your interest in contributing! This document describes the coding standards and conventions.

## Development Setup

```bash
pnpm install          # Install dependencies
pnpm build            # Build all packages
cd packages/<pkg> && pnpm build   # Build one package
cd packages/<pkg> && vp test run  # Run tests for one package
pnpm check            # Lint all packages
```

## Project Structure

```
packages/
  core/   — @office-open/core (descriptor runtime, DrawingML, chart/smartart, OPC)
  xml/    — @office-open/xml (XML parsing/serialization)
  docx/   — @office-open/docx (DOCX)
  pptx/   — @office-open/pptx (PPTX)
  xlsx/   — @office-open/xlsx (XLSX)
ooxml-schemas/  — OOXML XSD schemas (golden source of truth)
```

Every format package (docx, pptx, xlsx) follows the same layout:

```
src/
  parts/    — One module per OOXML XML part (types + descriptor co-located)
  shared/   — Types used by 2+ parts
  compiler.ts    — compileDocument/Presentation/Workbook()
  context.ts     — XxxWriteContext + XxxReadContext
  generate.ts    — generateDocument/Presentation/Workbook() entry
  parse.ts       — parseDocument/Presentation/Workbook() entry
  patch.ts       — patchDocument/Presentation/Workbook() entry
                   (single file when simple; use patch/ directory if complex)
  util/          — Helpers
  index.ts       — Public API
```

## OOXML Schemas

`ooxml-schemas/` contains XSD schemas — the golden source of truth. Always reference these when implementing XML elements.

- `transitional/` — Used by all major software (primary reference)
- `strict/` — ISO/IEC 29500 standard
- `microsoft/` — Microsoft extensions

Key files: `wml.xsd` (DOCX), `pml.xsd` (PPTX), `sml.xsd` (XLSX), `dml-main.xsd` (DrawingML).

## Naming Conventions

### Files and Directories

Use **kebab-case** for all file and directory names.

```
parts/settings.ts           — simple part (single file)
parts/document/body.ts      — complex part (directory)
shared/run.ts               — shared types
```

### Descriptors

Each OOXML part has a `<part>Desc` descriptor with `stringify()` and `parse()`:

```typescript
export const settingsDesc: CustomDescriptor<SettingsOptions> = {
  kind: "custom",
  stringify(opts, ctx) {
    return xml;
  },
  parse(el, ctx) {
    return opts;
  },
};
```

### Interfaces

**PascalCase** without `I` prefix. Configuration interfaces use `Options` suffix. All properties `readonly`.

```typescript
export interface ParagraphOptions {
  readonly alignment?: string;
  readonly children?: readonly (RunOptions | string)[];
}
```

### Functions

Use **camelCase**. Follow the appropriate prefix convention:

| Prefix       | Purpose                              | Example                                            |
| ------------ | ------------------------------------ | -------------------------------------------------- |
| `stringify*` | Generate XML from Options            | `stringifyRunProperties()`, `stringifyParagraph()` |
| `parse*`     | Parse XML into Options               | `parseBody()`, `parseRun()`                        |
| `create*`    | Factory functions for XML elements   | `createOutline()`, `createBevel()`                 |
| `build*`     | Build lookup tables or composite XML | `buildContentTypes()`, `buildTransition()`         |
| `compile*`   | Top-level compilation entry          | `compileDocument()`, `compilePresentation()`       |

### Constants (Enumerated Types)

Use `as const` objects (not TypeScript `enum`). Keys use **SCREAMING_SNAKE_CASE**. Values use **lowercase full English words**.

```typescript
export const AlignmentType = {
  START: "start", // XSD: "start"
  CENTER: "center", // XSD: "center"
} as const;

// When XSD uses abbreviations — map to full words
export const TextAlignment = {
  LEFT: "left", // XSD: "l"
  CENTER: "center", // XSD: "ctr"
} as const;
```

### Property Naming

Interface properties use **full English words** (camelCase), even when XML uses abbreviations:

```typescript
outline      → a:ln       gradientFill → a:gradFill
outerShadow  → a:outerShdw  solidFill    → a:solidFill
```

## Options Interface Design

### Flat vs Nested

| Pattern    | When to use                                      | Example                                     |
| ---------- | ------------------------------------------------ | ------------------------------------------- |
| **Flat**   | Simple, independent properties                   | `{ alignment, spacing, indent }`            |
| **Nested** | Properties map to a single XSD container element | `{ borders: { top, bottom, left, right } }` |

Rule: if 3+ properties share the same prefix, nest them under a property that names the concept and matches the XSD container.

### Container Field Naming

| Pattern       | Field Name  | When to use         | Example                   |
| ------------- | ----------- | ------------------- | ------------------------- |
| Heterogeneous | `children`  | Mixed element types | `SectionOptions.children` |
| Homogeneous   | Domain name | Single element type | `TableOptions.rows`       |

Domain names follow the XSD element: `rows` for `w:tr`/`x:row`, `cells` for `w:tc`/`x:c`.

## Descriptor Pattern

All XML serialization uses the descriptor pattern from `@office-open/core/descriptor`:

- **`CustomDescriptor<T>`** — for complex parts with custom stringify/parse logic
- **`ElementDescriptor<T>`** — for simple declarative attr/child mapping
- **`element<T>(tag)`** — builder for `ElementDescriptor`

Each descriptor is **bidirectional**: has both `stringify()` and `parse()`.

## XML Generation

XML is generated via **string concatenation** (template literals), not intermediate object trees. For complex dynamic XML, use `buildXml()` (re-export of `element()` from `@office-open/xml`).

```typescript
// Simple — inline template
const xml = `<a:noFill/>`;
const xml = `<a:off x="${x}" y="${y}/>`;

// Dynamic — array push + join
const parts: string[] = [];
if (opts.fill) parts.push(stringifyFill(opts.fill));
return `<p:spPr>${parts.join("")}</p:spPr>`;
```

## Loop Patterns

| Scenario                            | Use                 | Reason                            |
| ----------------------------------- | ------------------- | --------------------------------- |
| Transform into new array            | `.map()`            | Expresses "transform" intent      |
| Filter elements                     | `.filter()`         | Expresses "filter" intent         |
| Side-effect iteration, async, break | `for...of`          | Full control, supports early exit |
| Performance-sensitive hot paths     | `for...of` or `for` | ~3x faster than `.forEach()`      |

**Avoid `.forEach()`** — `for...of` is strictly superior.

## XSD Value Mapping

When XSD uses abbreviations, mapping is centralized in `packages/core/src/xsd-mappings.ts`. The mapping is bidirectional:

- **Generation** (Options → XML): user-friendly → XSD abbreviated
- **Parsing** (XML → Options): XSD abbreviated → user-friendly

When XSD uses full words (e.g. `"center"`), no mapping needed.

## Running Demos

```bash
cd packages/docx && pnpm tsx demo/<demo-file>.ts
cd packages/pptx && pnpm tsx demo/<demo-file>.ts
cd packages/xlsx && pnpm tsx demo/<demo-file>.ts
```

## Validation

```bash
pnpm tsx scripts/validate.ts                # All demos
pnpm tsx scripts/validate.ts pptx           # One package
pnpm tsx scripts/validate.ts docx "path.docx"  # Specific file
```

## Pull Request Process

1. `pnpm check` passes with no errors
2. Run relevant demos and tests to verify changes
3. For new XML elements, validate output against XSD schemas
4. Follow naming conventions described above
5. Keep changes minimal and focused — match existing style
