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
parts/document/body.ts      — complex part (subdirectory)
shared/shape/shape.ts       — graphic-object types AND descriptor, co-located
```

Two co-location patterns: **docx/xlsx** put each part's types and `<part>Desc` descriptor together in `parts/<part>/` (complex) or `parts/<part>.ts` (simple); **pptx** keeps descriptors in a dedicated `parts/descriptors/` layer with public types in `shared/<domain>/`. Both are acceptable — pick one per package and stay consistent within it. Complex parts are subdirectories: xlsx `worksheet`/`workbook`/`styles`/`pivot-table`/`revision-log`/`drawing` are large enough to warrant subdirectories, matching the docx baseline.

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

**PascalCase** without `I` prefix. Configuration interfaces use `Options` suffix. Do **not** mark properties `readonly` — the parse path assigns into these objects.

```typescript
export interface ParagraphOptions {
  alignment?: string;
  children?: (RunOptions | string)[];
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

### Enumerated Types

Prefer **string literal unions** for enumerated option types. They are pure types (no runtime object), keep the call site self-documenting (`position: "b"`), and have no runtime cost. Map XSD abbreviations to full words in JSDoc, not via a value object.

```typescript
// Preferred: string literal union
export type AxisPosition = "b" | "l" | "r" | "t";

/** XSD ST_TickMark: cross / in / none / out */
export type AxisTickMark = "cross" | "in" | "none" | "out";
```

Use an `as const` object **only when the values are referenced at runtime** — iteration, `Record<Enum, T>` lookup keys, or `Enum.KEY` access in demos/tests. The runtime object must justify itself; a pure option-field type stays a string union. Keys then use **SCREAMING_SNAKE_CASE**, values **lowercase full English words**.

```typescript
// Justified: consumed as Record keys / iterated at runtime
const CHART_TYPE_TAGS: Record<ChartType, string> = { column: "c:barChart", ... };
```

Do **not** use TypeScript `enum`. Historical `as const` enumerated option types in docx/pptx/xlsx are migrated to string unions opportunistically when touched; the chart module is fully migrated.

### Property Naming

Interface properties use **full English words** (camelCase), even when XML uses abbreviations. Do **not** abbreviate by deleting letters within a word — write `index`, not `idx` (aligns with the Google TypeScript Style Guide). Reference elements map to `*Reference`, matching the Open XML SDK (`LineReference`, `FillReference`, …).

```typescript
// Element names → semantic full words
outline         → a:ln         gradientFill  → a:gradFill
outerShadow     → a:outerShdw  solidFill     → a:solidFill

// Reference elements → *Reference (Open XML SDK alignment)
lineReference   → a:lnRef      fillReference    → a:fillRef
effectReference → a:effectRef  fontReference    → a:fontRef
```

**OOXML standard attribute tokens are preserved verbatim** — a 1:1 mapping with the XSD attribute name, exactly like `@id` → `id`. Keep these as-is: `id`, `idx`, `numFmt`, `numId`, `fontId`, `fillId`, `borderId`, `clrIdx`, `lastIdx`.

**Never invent compound abbreviations** by concatenating an element abbreviation with an attribute. The `@idx` of `a:lnRef` is `lineReferenceIndex`, never `lnIdx` (`ln` + `idx`).

**Read/write symmetry:** the same OOXML concept uses the same property name on both the write `Options` and the parse output — `fontId` on both `CellXfEntry` and `IndexedXfEntry`, never `fontId` on one and `fontIdx` on the other.

### Cross-Package Naming

The same concept uses the **same name in every package** — picture/shape/connector/group share one `*Options` name across docx/pptx/xlsx (`PictureOptions` everywhere, not per-package `DrawingDescriptorOptions`/`PictureDescriptorOptions`/`DrawingImageOptions`).

- **No grouping prefixes** (`Model`/`Content`/`Element`/`Universal`). `UniversalMeasure` is the XSD `ST_UniversalMeasure` type, kept as-is — not a precedent.
- **Tables are the exception**: each package's `TableOptions` models a genuinely different thing (`w:tbl` flow / `a:tbl` graphic / xlsx cell-range), so the name stays but each carries a JSDoc boundary note; docx/pptx `extends` the shared `core/table/` `BaseTableOptions` (rows/cells/span/6-flags/columnWidths/vertical-align) and add domain-specific style/position, xlsx is independent ([Cross-Format Conversion](#cross-format-conversion)).
- **Renames**: pre-1.0, rename in place (no `@deprecated` aliases — `GroupOptions`/`ConnectorOptions`/`ShapeOptions` replaced `WpgGroupRunOptions`/`ConnectorShapeOptions`/`WpsShapeRunOptions` directly); post-1.0, keep the old name as a `@deprecated` alias until the next major release.

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

## Measurement Units

Geometry and sizing fields accept **`number`** (the format's native unit) or a **`UniversalMeasure` string** (mm/cm/in/pt/pc/pi; px at 96 DPI on DrawingML fields). The native unit follows the domain:

| Domain                                              | `number` means    | Convert with    |
| --------------------------------------------------- | ----------------- | --------------- |
| DrawingML geometry (shapes, images, charts, tables) | EMU               | `convertToEmu`  |
| Word spacing / indent / font size                   | twip / half-point | `convertToTwip` |
| xlsx row height                                     | points            | `convertToPt`   |
| xlsx page margins                                   | inches            | `convertToInch` |

The converters are **polymorphic** — a `number` passes through unchanged, a string is parsed — so a single field accepts both forms.

**Round-trip is lossless.** Parse returns the native unit verbatim (e.g. EMU), and stringify converts any string back to that unit. Never round-trip through pixels: it quantizes to the grid and is irreversible.

### Input convenience, XSD-valid output

`UniversalMeasure` exists for input ergonomics only — the XML we emit must still satisfy its XSD type, so stringify always converts to the integer/unit the schema requires, never writing a raw UM string where a number is mandated.

A field stays plain `number` when its value isn't a geometric length or its XSD type is integer-only — bevel size, 3D angles, rotation, xlsx column width (character units). Don't add `UniversalMeasure` to such fields.

## Descriptor Pattern

All XML serialization uses the descriptor pattern from `@office-open/core/descriptor`:

- **`CustomDescriptor<T>`** — every descriptor is custom: hand-written `stringify()` + `parse()` for the part

Each descriptor is **bidirectional**: has both `stringify()` and `parse()`.

## Cross-Format Conversion

Cross-format copy works at the `Options` layer — **no unified document model**.

- **Similar structures** (picture/connector/group): each concept has a core base the format `Options` extend — `core/picture/` `BasePictureOptions`, `core/connector/` `BaseConnectorOptions`, `core/group/` `BaseGroupOptions` (all `extends NonVisualDrawingPropertiesOptions`, the shared `a:CT_NonVisualDrawingProps` cNvPr/docPr type from `core/drawingml/non-visual/`); docx picture/group stay a discriminated union bridged through `altText`. `convert/*.ts` passes the base through directly (cNvPr name/description/title/hidden threads every leg) and maps only container/positioning (`wps:`/`a:sp`/`xdr:sp`; inline vs x/y vs cell anchor) plus format-specific spPr convenience. Shape has no base (YAGNI — cNvPr already unified, spPr/textBody already in core, no shape-specific field worth lifting).
- **Tables**: three packages model tables fundamentally differently, but they share a structural core (rows/cells/span/6-flags/columnWidths/vertical-align) in `core/table/` (`BaseTableOptions`/`BaseTableRowOptions`/`BaseTableCellOptions`). docx/pptx `TableOptions extends Base*` and add domain-specific style/position; xlsx is independent (sml Table is a data range). `convert/table.ts` passes the base through directly and translates only cell content (w:p↔a:p via `convert/text`), units (twip↔EMU via `core/util/converters`), and pptx fill/scheme-color → docx shading/themeColor.
- **Text**: `a:p` is shared in `core/drawingml/text/`; docx bridges `a:p ↔ w:p` (font ×100↔×2 half-points, `srgbClr`↔`w:color`, typeface↔`rFonts`, hyperlinks).
- **Charts**: unified via `core/chart/`; cross-format is only packaging.

Conversion functions are per-package pure functions (`to*`/`convert*`/`from*`), exported from `@office-open/<pkg>` — never core, to keep `format → core` one-directional. **Fidelity** matches MS Office paste: shapes (spPr/fill/outline/effects/textBody) translate near-losslessly across pptx/xlsx/docx; positioning is heuristic (xlsx cell-anchor ↔ absolute EMU), lossy on xlsx hops; connectors skip docx (no standalone connector element); table structure/merge/border/font/alignment high-fidelity; multi-paragraph text ↔ single cell, and formulas/validation, lossy.

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

When XSD uses abbreviations, mapping is centralized in `packages/core/src/util/mappings.ts`. Each mapping is a `bidi()` helper exposing `.to()` (user value → XSD value) and `.from()` (XSD value → user value). The mapping is bidirectional:

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
