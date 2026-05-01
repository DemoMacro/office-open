You are a senior TypeScript developer

## Project

**office-open** is a monorepo for generating Office Open XML documents (.docx, .pptx, .xlsx) with JS/TS.

## Structure

- `packages/docx` - @office-open/docx (main DOCX package)
- `packages/docx-plus` - docx-plus (compat re-export)
- `packages/core` - @office-open/core (shared infrastructure, WIP)
- `packages/xml` - @office-open/xml (WIP)
- `packages/pptx` - @office-open/pptx (WIP)
- `packages/xlsx` - @office-open/xlsx (WIP)

## OOXML Specification

The `ooxml-schemas/` directory contains the official ISO-IEC29500 OOXML XSD schemas. These are the **golden source of truth** for all OOXML element names, attributes, and structure.

Key schemas:

- `wml.xsd` - WordprocessingML (documents)
- `pml.xsd` - PresentationML (presentations)
- `sml.xsd` - SpreadsheetML (spreadsheets)
- `dml-main.xsd` - DrawingML (images, shapes)
- `shared-math.xsd` - Math equations

## Code Style

- TypeScript with strict mode
- Path aliases in packages/docx: `@file/`, `@export/`, `@util/`
- Classes extend `XmlComponent` for XML elements

## Build & Test

- **Packaging**: `basis build`
- **Testing**: `vp test run --coverage` (Vitest via vite-plus)

## Running Demos

```bash
pnpm run-ts -- ./packages/docx/demo/<demo-file>.ts
```
