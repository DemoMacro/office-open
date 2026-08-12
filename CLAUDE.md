You are a senior TypeScript developer.

## Project

**office-open** is a monorepo for generating and parsing Office Open XML documents (.docx, .pptx, .xlsx) with JS/TS. Declarative API, works in Node.js and browsers. Bidirectional: stringify (JSON → XML) and parse (XML → JSON).

## Architecture

- **Packages**: `core/` (shared OOXML domains — descriptor runtime, `drawingml/`, `chart/`, smartart, OPC), `xml/`, `docx/`, `pptx/`, `xlsx/`. The three format packages are peers: same `src/` layout (`parts/`, `shared/`, `compiler.ts`, `context.ts`, `generate.ts`, `parse.ts`, `patch.ts`, `index.ts`) and one `*Options` name per shared concept.
- **Parts**: one module per OOXML XML part. docx/xlsx co-locate types and the `<part>Desc` descriptor; pptx keeps descriptors in `parts/descriptors/` with public types in `shared/<domain>/`. Cross-part shared types live in `shared/`.
- **Descriptor pattern**: every part is a `CustomDescriptor<T>` with hand-written `stringify(opts, ctx)` + `parse(el, ctx)` (bidirectional). Runtime at `packages/core/src/descriptor/`.
- **Cross-format copy** reuses `core`'s shared domains + package-to-package conversion in `packages/office-open/src/convert/` — **no unified document-model layer**. The conversion mode follows how the concept's XML maps across packages:
  - _Identical XML_ (`chart`: one `c:chartSpace` part) → shared core model, no conversion.
  - _Same element, different anchoring_ (`picture` `pic:pic`/`p:pic`/`xdr:pic`; shapes; text) → each package owns an `*Options` reflecting its anchor model (docx `wp:inline`/`wp:anchor`, pptx spTree, xlsx cell anchor) + pairwise `convert/*` functions; N is small, so pairwise beats a hub-spoke intermediate.
  - _Structurally different XML_ (`table`: `w:tbl` flow / `a:tbl` graphic / `sml` cell-range) → target state is a prefix-free core intermediate (`core/table/` `TableGrid`) + per-package `from<Pkg>`/`to<Pkg>` adapters, mirroring `core/chart/`; **not yet implemented**.
- **OOXML XSD** (`ooxml-schemas/transitional/`) is the golden source of truth — `wml.xsd` (DOCX), `pml.xsd` (PPTX), `sml.xsd` (XLSX), `dml-main.xsd` (DrawingML). Validate XML output against it.

Full layout and conventions in [CONTRIBUTING.md](./CONTRIBUTING.md).

## Build & Test

- **Build**: `pnpm build` (all) or `cd packages/<pkg> && pnpm build` (one)
- **Test**: `cd packages/<pkg> && pnpm exec vp test run` (`vp` is not on PATH)
- **Lint**: `pnpm check` (resolves via `dist/` — build first)
- **Validate (XSD)**: `pnpm tsx scripts/validate.ts`

## Measurement Units

Geometry/sizing fields take **`number`** (native unit — EMU for DrawingML, twip for Word) or a **`UniversalMeasure` string** (mm/cm/in/pt/pc/pi, plus px at 96 DPI on DrawingML). Convert with `convertToEmu` / `convertToTwip` / `convertToPt` / `convertToInch` — polymorphic, `number` passes through, round-trip lossless. `UniversalMeasure` is input-only: stringify emits the integer/unit the XSD requires, so a field stays plain `number` when it isn't a geometric length or its XSD type is integer-only (bevel, 3D, rotation). See [CONTRIBUTING.md](./CONTRIBUTING.md#measurement-units).

## Code Conventions

Project-specific core; full standards in [CONTRIBUTING.md](./CONTRIBUTING.md).

- **Naming**: `<part>Desc` descriptors, `<Part>Options` interfaces, `stringify*()` / `parse*()` / `patch*()` helpers. One concept, one `*Options` name across packages (`PictureOptions` in docx/pptx/xlsx). No `Model`/`Content`/`Element` prefixes (`UniversalMeasure` is XSD `ST_UniversalMeasure`, not a precedent). Public-API renames use `@deprecated` aliases.
- **OOXML-fidelity naming**: names mirror the OOXML element they serialize, not library legacy. A picture frame (`pic:pic`/`p:pic`/`xdr:pic`) is `picture` everywhere (`PictureOptions`, `{ picture }`, `picture-run`, `createPictureData`); `image` survives only where OOXML uses it — MIME strings (`image/png`), OPC relationship URIs (`…/relationships/image`), `a:blip`/VML bitmap data, the page background (`BackgroundImageOptions`). Internal symbols take a clean break, not `@deprecated`.
- **Properties**: OOXML attribute tokens verbatim (`id`/`idx`/`numFmt`/`fontId`); reference elements → `*Reference`; never compound abbreviations (`lnIdx` → `lineReferenceIndex`); full English words otherwise.
- **Enums**: string literal unions by default; `as const` objects only when values are referenced at runtime (point access, `Record` keys, `Object.values`).
- **XML generation**: template-literal concatenation — no intermediate object trees.
- Convention changes (structure / naming / cross-format) land here and in CONTRIBUTING.md **before** code — docs are the target-state spec.
