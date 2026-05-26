You are a senior TypeScript developer.

## Project

**office-open** is a monorepo for generating Office Open XML documents (.docx, .pptx, .xlsx) with JS/TS. Declarative API, works in Node.js and browsers.

## Structure

- `packages/core` - @office-open/core (shared XML components, formatter, chart/smartart, unit converters)
- `packages/xml` - @office-open/xml (XML parsing/serialization, replacing xml + xml-js)
- `packages/docx` - @office-open/docx (main DOCX package)
- `packages/docx-plus` - docx-plus (compat re-export of @office-open/docx)
- `packages/xlsx` - @office-open/xlsx (TODO)
- `packages/pptx` - @office-open/pptx (PPTX generation package)
- `ooxml-schemas/` - ISO-IEC29500 OOXML XSD schemas

## Build & Test

- **Packaging**: `basis build` (per package) or `pnpm build` (all packages)
- **Testing**: `vp test run` (vite-plus test runner, per package)
- **Lint**: `pnpm check` (from root, runs vp check across all packages)

## OOXML Specification

`ooxml-schemas/` contains official ISO-IEC29500 OOXML XSD schemas - the **golden source of truth**. Always reference these when implementing XML elements.

- `wml.xsd` - WordprocessingML (documents)
- `pml.xsd` - PresentationML (presentations)
- `sml.xsd` - SpreadsheetML (spreadsheets)
- `dml-main.xsd` - DrawingML (images/shapes)
- `shared-math.xsd` - Math

## Code Conventions

- Path aliases in packages/docx: `@file/`, `@export/`, `@util/`
- Classes extend `XmlComponent` for XML elements
- Use `Formatter` to convert components to XML tree

### Interface Naming

Do **not** use the `I` prefix for interfaces. This aligns with modern TypeScript conventions (TypeScript official, Google, ts.dev style guides).

- e.g., `PropertiesOptions`, `SectionOptions`, `ParagraphOptions`, `TableOptions`, `ImageOptions`
- e.g., `OutlineOptions`, `SolidFillOptions`, `GradientFillOptions`, `ColorOptions`, `EffectListOptions`
- e.g., `MediaData`, `MediaDataTransformation`, `Context`, `XmlableObject`, `GradientStop`

### Property Naming vs XML Element Names

Interface property names use **full English words**, even when the corresponding XML element uses XSD abbreviations:

- `outline` → `a:ln`, `gradientFill` → `a:gradFill`, `outerShadow` → `a:outerShdw`

Only use XSD abbreviations when the abbreviation IS the standard English term (e.g., `solidFill`, `noFill`, `glow`).

**Best practices:**

- Verify XML output structure matches OOXML spec
- Test option combinations and edge cases
- Descriptive test names explaining behavior

## Behavioral Guidelines

### Think Before Coding

- State assumptions explicitly. If uncertain, ask before implementing.
- If multiple interpretations exist, present them — don't pick silently.
- If a simpler approach exists, say so. Push back when warranted.

### Simplicity First

- No features beyond what was asked. No speculative abstractions.
- If 200 lines could be 50, rewrite it.

### Surgical Changes

- Touch only what you must. Match existing style.
- Don't "improve" adjacent code. Don't refactor things that aren't broken.
- Remove only orphans that YOUR changes created, not pre-existing dead code.

### Goal-Driven Execution

- Transform tasks into verifiable goals. State a brief plan for multi-step tasks.
- Loop until verified — don't wait for clarification after obvious failures.

## Running Demos

```bash
# Docx demos (run from packages/docx)
cd packages/docx && pnpm tsx demo/<demo-file>.ts

# PPTX demos (run from packages/pptx)
cd packages/pptx && pnpm tsx demo/<demo-file>.ts
```

## Validation

Validate generated XML against OOXML XSD schemas:

```bash
# Validate all demos (pptx + docx)
pnpm tsx scripts/validate.ts

# Validate single package
pnpm tsx scripts/validate.ts pptx
pnpm tsx scripts/validate.ts docx

# Validate specific demo file
pnpm tsx scripts/validate.ts slide "packages/pptx/My Presentation.pptx" [slide-number]
pnpm tsx scripts/validate.ts docx "packages/docx/My Document.docx"
```
