# Contributing to office-open

Thank you for your interest in contributing! This document describes the coding standards and conventions used in this project.

## Development Setup

```bash
# Install dependencies
pnpm install

# Build all packages
pnpm build

# Build a specific package
cd packages/docx && pnpm build

# Run tests for a specific package
cd packages/docx && vp test run

# Lint all packages
pnpm check
```

## Project Structure

```
packages/
  core/   - @office-open/core (shared XML components, formatter, chart/smartart, unit converters)
  xml/    - @office-open/xml (XML parsing/serialization)
  docx/   - @office-open/docx (DOCX generation)
  pptx/   - @office-open/pptx (PPTX generation)
  xlsx/   - @office-open/xlsx (XLSX generation)

ooxml-schemas/  - OOXML XSD schemas (the golden source of truth)
```

## OOXML Specification

The `ooxml-schemas/` directory contains OOXML XSD schemas. Always reference these when implementing XML elements.

- `transitional/` - Transitional OOXML schemas (used by all major software)
- `strict/` - ISO/IEC 29500 standard schemas
- `microsoft/` - Microsoft extension schemas

Key schema files:

- `wml.xsd` - WordprocessingML (documents)
- `pml.xsd` - PresentationML (presentations)
- `dml-main.xsd` - DrawingML (images/shapes)

## Naming Conventions

### Files and Directories

Use **kebab-case** for all file and directory names.

```
file/paragraph/formatting/alignment.ts
file/drawing/floating/floating-position.ts
parse/numbering.ts
```

### Classes

Use **PascalCase** for class names. Classes extend `XmlComponent` for XML element components.

```typescript
export class Paragraph extends XmlComponent { ... }
export class Shape extends Xc { ... }
```

### Interfaces

Use **PascalCase** without the `I` prefix. Configuration interfaces use the `Options` suffix.

```typescript
export interface ParagraphOptions { ... }
export interface ShapeOptions { ... }
export interface BorderOptions { ... }
```

All interface properties should use the `readonly` modifier.

### Functions

Use **camelCase** for function names. Follow the appropriate prefix convention:

| Prefix    | Purpose                                      | Example                                       |
| --------- | -------------------------------------------- | --------------------------------------------- |
| `create*` | Factory functions that build XML elements    | `createBevel()`, `createShape3D()`            |
| `parse*`  | Functions that parse XML into Options        | `parseSlide()`, `parseNumberingDefinitions()` |
| `build*`  | Functions that build lookup tables or caches | `buildStyleCache()`, `buildNumberingCache()`  |

### Constants (Enumerated Types)

Use `as const` objects (not TypeScript `enum`). Keys use **SCREAMING_SNAKE_CASE**. Values use **lowercase full English words**.

```typescript
// When XSD uses full words - use them directly (no mapping needed)
export const AlignmentType = {
  START: "start", // XSD: "start"
  CENTER: "center", // XSD: "center"
  END: "end", // XSD: "end"
} as const;

// When XSD uses abbreviations - map to full words
export const TextAlignment = {
  LEFT: "left", // XSD: "l"
  CENTER: "center", // XSD: "ctr"
  RIGHT: "right", // XSD: "r"
  JUSTIFY: "justify", // XSD: "just"
} as const;
```

Type references use `keyof typeof` or `ValueOf<typeof>`:

```typescript
type Alignment = keyof typeof AlignmentType;
type AlignmentValue = (typeof AlignmentType)[keyof typeof AlignmentType];
```

### Property Naming

Interface property names use **full English words** (camelCase), even when the corresponding XML element uses abbreviations:

```typescript
// Property name → XML element
outline      → a:ln
gradientFill → a:gradFill
outerShadow  → a:outerShdw
solidFill    → a:solidFill
```

### Getters and Setters

Use **camelCase** for class getters and setters (consistent with method naming):

```typescript
// Good
public get shapeId(): number { ... }

// Avoid
public get ShapeId(): number { ... }
```

## Options Interface Design

### When to Use Flat vs Nested

| Pattern    | When to use                                                                           | Example                                     |
| ---------- | ------------------------------------------------------------------------------------- | ------------------------------------------- |
| **Flat**   | Simple, independent properties                                                        | `{ alignment, spacing, indent }`            |
| **Nested** | Properties form a cohesive domain concept that maps to a single XSD container element | `{ borders: { top, bottom, left, right } }` |

Rule of thumb: if 3+ properties share the same prefix (e.g. `textVertical`, `textAnchor`, `textWrap`), nest them under a single property that names the concept and matches the XSD container element (e.g. `textBody: { vertical, anchor, wrap }` for `p:txBody` > `a:bodyPr`).

This aligns with the XSD structure where related attributes are grouped under a single container element (e.g. `a:bodyPr` contains `vert`, `anchor`, `wrap`).

### Parameters: Positional vs Options Object

| Pattern                   | When to use                                                      |
| ------------------------- | ---------------------------------------------------------------- |
| **Positional parameters** | All parameters are required and stable (e.g. `attr(el, name)`)   |
| **Options object**        | 2+ optional parameters, interface may evolve, public API surface |

For internal helper functions with all-required parameters, prefer positional parameters even if there are more than 3.

## Loop Patterns

| Scenario                                                    | Use                 | Reason                                          |
| ----------------------------------------------------------- | ------------------- | ----------------------------------------------- |
| Transform data into a new array                             | `.map()`            | Expresses "transform" intent, returns new array |
| Filter elements                                             | `.filter()`         | Expresses "filter" intent                       |
| Side-effect iteration, async/await, need `break`/`continue` | `for...of`          | Full control, supports early exit and async     |
| Performance-sensitive hot paths                             | `for...of` or `for` | ~3x faster than `.forEach()`                    |

**Avoid `.forEach()`** - `for...of` is a strictly superior replacement (more flexible, faster, supports `async/await` and `break`/`continue`).

## XSD Value Mapping

When XSD uses abbreviations that don't match human-readable values, a mapping layer is used. All mappings are centralized in `packages/core/src/xsd-mappings.ts`.

The mapping is bidirectional:

- **Generation** (Options → XML): user-friendly values → XSD abbreviated values
- **Parsing** (XML → Options): XSD abbreviated values → user-friendly values

When XSD already uses full English words (e.g. `"center"`, `"start"`), no mapping is needed - use the XSD value directly.

## Running Demos

```bash
# Docx demos (run from packages/docx)
cd packages/docx && pnpm tsx demo/<demo-file>.ts

# Pptx demos (run from packages/pptx)
cd packages/pptx && pnpm tsx demo/<demo-file>.ts

# Xlsx demos (run from packages/xlsx)
cd packages/xlsx && pnpm tsx demo/<demo-file>.ts
```

## Validation

Validate generated XML against OOXML XSD schemas:

```bash
# Validate all demos
pnpm tsx scripts/validate.ts

# Validate single package
pnpm tsx scripts/validate.ts pptx
pnpm tsx scripts/validate.ts docx
pnpm tsx scripts/validate.ts xlsx

# Validate specific file
pnpm tsx scripts/validate.ts docx "packages/docx/My Document.docx"
pnpm tsx scripts/validate.ts xlsx "packages/xlsx/My Workbook.xlsx"
```

## Pull Request Process

1. Ensure `pnpm check` passes with no errors
2. Run relevant demos and tests to verify changes
3. For new XML elements, validate output against XSD schemas
4. Follow the naming conventions described above
5. Keep changes minimal and focused - match existing style
