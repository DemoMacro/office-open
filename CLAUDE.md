You are a senior TypeScript developer.

## Project

**office-open** is a monorepo for generating and parsing Office Open XML documents (.docx, .pptx, .xlsx) with JS/TS. Declarative API, works in Node.js and browsers. Bidirectional: stringify (JSON → XML) and parse (XML → JSON).

## Architecture

- **Packages**: `core/` (shared OOXML domains — descriptor runtime, `drawingml/`, `chart/`, smartart, OPC), `xml/`, `docx/`, `pptx/`, `xlsx/`. The three format packages are peers: same `src/` layout (`parts/`, `shared/`, `compiler.ts`, `context.ts`, `generate.ts`, `parse.ts`, `patch.ts`, `index.ts`) and same `*Options` names for shared concepts. Cross-format copy reuses `core`'s shared domains + package-to-package conversion functions — **no unified document-model layer**; where a concept's XML is structurally broken across packages (tables), it owns a prefix-free intermediate in its own core domain (`core/table/` `TableGrid`), mirroring `core/chart/`. Full layout and conventions in [CONTRIBUTING.md](./CONTRIBUTING.md).
- **Parts**: one module per OOXML XML part; types and the `<part>Desc` descriptor are co-located. Simple parts are a single file (`parts/settings.ts`), complex parts are a directory (`parts/document/`). Cross-part shared types live in `shared/`.
- **Descriptor pattern**: every part is a `CustomDescriptor<T>` with hand-written `stringify(opts, ctx)` + `parse(el, ctx)` (bidirectional). Runtime at `packages/core/src/descriptor/`.
- **OOXML XSD** (`ooxml-schemas/transitional/`) is the golden source of truth — `wml.xsd` (DOCX), `pml.xsd` (PPTX), `sml.xsd` (XLSX), `dml-main.xsd` (DrawingML). Always validate XML output against it.

## Build & Test

- **Build**: `pnpm build` (all) or `cd packages/<pkg> && pnpm build` (one)
- **Test**: `cd packages/<pkg> && vp test run`
- **Lint**: `pnpm check` (runs `vp check` across all packages; resolves via `dist/`, so build first)
- **Validate (XSD)**: `pnpm tsx scripts/validate.ts`

## Measurement Units

Geometry/sizing fields take **`number`** (the format's native unit — EMU for DrawingML, twip for Word) or a **`UniversalMeasure` string** (mm/cm/in/pt/pc/pi, plus px at 96 DPI on DrawingML). Convert with `convertToEmu` / `convertToTwip` / `convertToPt` / `convertToInch` — polymorphic, `number` passes through. Round-trip is lossless (parse returns the native unit).

**XML output is always XSD-valid**: `UniversalMeasure` is an input convenience only — stringify emits the integer/unit the schema requires. A field stays plain `number` when it isn't a geometric length or its XSD type is integer-only (bevel, 3D, rotation). See [CONTRIBUTING.md](./CONTRIBUTING.md#measurement-units).

## Code Conventions

Full standards in [CONTRIBUTING.md](./CONTRIBUTING.md). Quick reference:

- **Naming**: `<part>Desc` descriptors, `<Part>Options` interfaces, `stringify*()` / `parse*()` / `patch*()` helpers. The same concept shares one `*Options` name across packages (`PictureOptions` in docx/pptx/xlsx); no `Model`/`Content`/`Element` prefixes (`UniversalMeasure` is XSD `ST_UniversalMeasure`, not a precedent); renames use `@deprecated` aliases — see [CONTRIBUTING.md#cross-package-naming](./CONTRIBUTING.md#cross-package-naming)
- **Properties**: full English words (camelCase); OOXML attribute tokens (`id`/`idx`/`numFmt`/`fontId`/…) preserved verbatim; reference elements → `*Reference`; never compound abbreviations like `lnIdx` → `lineReferenceIndex` — see [CONTRIBUTING.md#property-naming](./CONTRIBUTING.md#property-naming)
- **Enums**: string literal unions by default; `as const` objects (SCREAMING_SNAKE_CASE keys, lowercase values) only when values are referenced at runtime
- **Files**: kebab-case, no `I` prefix on interfaces, no `readonly` on Options properties
- **Loops**: `for...of` default, `.map()` only when returning a new array
- **XML generation**: string concatenation via template literals, no intermediate object trees

## Behavioral Guidelines

- State assumptions explicitly. If uncertain, ask before implementing.
- No features beyond what was asked. No speculative abstractions.
- Touch only what you must. Match existing style.
- Transform tasks into verifiable goals. Loop until verified.
- Convention changes (structure/naming/cross-format) land in CLAUDE.md and CONTRIBUTING.md **before** code — docs are the target-state spec.
