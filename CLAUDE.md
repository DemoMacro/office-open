You are a senior TypeScript developer.

## Project

**office-open** is a monorepo for generating and parsing Office Open XML documents (.docx, .pptx, .xlsx) with JS/TS. Declarative API, works in Node.js and browsers. Bidirectional: stringify (JSON → XML) and parse (XML → JSON).

## Architecture

- **Packages**: `core/` (shared OOXML domains — descriptor runtime, `drawing/`, `chart/`, `table/`, `picture/`, `connector/`, `group/`, smartart, OPC), `xml/`, `docx/`, `pptx/`, `xlsx/`. The three format packages are peers: same `src/` layout (`parts/`, `shared/`, `compiler.ts`, `context.ts`, `generate.ts`, `parse.ts`, `patch.ts`, `index.ts`) and one `**Options` name per shared concept.
- **Parts**: one module per OOXML XML part. docx/xlsx co-locate types and the `<part>Desc` descriptor; pptx keeps descriptors in `parts/descriptors/` with public types in `shared/<domain>/`. Cross-part shared types live in `shared/`.
- **Descriptor pattern**: every part is a `CustomDescriptor<T>` with hand-written `stringify(opts, ctx)` + `parse(el, ctx)` (bidirectional). Runtime at `packages/core/src/descriptor/`.
- **Cross-format copy** reuses `core`'s shared domains + package-to-package conversion in `packages/office-open/src/convert/` — **no unified document-model layer**. The conversion mode follows how the concept's XML maps across packages:
  - _Identical XML_ (`chart`: one `c:chartSpace` part) → shared core model, no conversion.
  - _Same element, different anchoring_ (`picture` `pic:pic`/`p:pic`/`xdr:pic`; shapes; text) → each package owns an `*Options` reflecting its anchor model (docx `wp:inline`/`wp:anchor`, pptx spTree, xlsx cell anchor) + pairwise `convert/*` functions; N is small, so pairwise beats a hub-spoke intermediate. The shared payload still lands in core: `core/picture/` `BasePictureOptions` (data/type + cNvPr via `NonVisualDrawingPropertiesOptions`) is extended by pptx/xlsx `PictureOptions`, while docx stays a format discriminated union bridged through its `altText`; `convert/picture.ts` threads the cNvPr fields (name/description/title/hidden) straight through every leg. The same `NonVisualDrawingPropertiesOptions` (a:CT_NonVisualDrawingProps) backs every drawing's cNvPr/docPr across all three packages via `core/drawing/non-visual/`.
  - _Structurally different XML_ (`table`: `w:tbl` flow / `a:tbl` graphic / `sml` cell-range) → `core/table/` defines the shared structural base (`BaseTableOptions`/`BaseTableRowOptions`/`BaseTableCellOptions`: rows/cells/span/6-flags/columnWidths/vertical-align); docx/pptx `TableOptions extends Base*` add domain-specific style/position, xlsx is independent (its sml Table is a data range, restored visually via `convert/table.ts`). `convert/table.ts` passes the structural base through directly and translates only cell content (w:p↔a:p via `convert/text`), units (twip↔EMU via `core/util/converters`), and pptx fill/scheme-color → docx shading/themeColor.
  - _Connector_ (`cxnSp`: `p:cxnSp`/`xdr:cxnSp`; docx has no standalone cxnSp) → `core/connector/` `BaseConnectorOptions extends NonVisualDrawingPropertiesOptions` (cNvPr + locking + start/end connection endpoints; no spPr — pptx top-level convenience vs xlsx spPr asymmetry, kept per-package). pptx `ConnectorOptions`/`LineShapeOptions` and xlsx `ConnectorOptions` extend base; `convert/connector.ts` threads cNvPr + locking + endpoints via `pickConnectorBase`.
  - _Group_ (`grpSp`: `p:grpSp`/`xdr:grpSp`/`wpg:wgp`) → `core/drawing/group-shape-properties-desc.ts` `GroupShapePropertiesOptions` (CT_GroupShapeProperties: xfrm/fill/effects/scene3d — no geometry/ln/sp3d, distinct from CT_ShapeProperties) + `core/group/` `BaseGroupOptions extends NonVisualDrawingPropertiesOptions` (cNvPr only; grpSpPr content/children/position stay per-package). pptx/xlsx/docx all use `GroupOptions` (docx bridged through `altText` like picture). `convert/group.ts` threads container + child cNvPr via `pickGroupBase`/`pickNonVisualDrawingProperties`.
  - _Shape base is YAGNI_: cNvPr is already unified via `NonVisualDrawingPropertiesOptions` (Phase 2), spPr/textBody already live in core, there are no scattered convert adapters and no shape-specific public field worth lifting — so no `BaseShapeOptions` is introduced.
- **OOXML XSD** (`ooxml-schemas/transitional/`) is the golden source of truth — `wml.xsd` (DOCX), `pml.xsd` (PPTX), `sml.xsd` (XLSX), `dml-main.xsd` (DrawingML). Validate XML output against it.

Full layout and conventions in [CONTRIBUTING.md](./CONTRIBUTING.md).

## Build & Test

- **Build**: `pnpm build` (all) or `cd packages/<pkg> && pnpm build` (one)
- **Test**: `cd packages/<pkg> && pnpm exec vp test run` (`vp` is not on PATH)
- **Lint**: `pnpm check` (resolves via `dist/` — build first)
- **Validate (XSD)**: `pnpm tsx scripts/validate.ts` — runs every package demo (parallel pool, outputs to `packages/<pkg>/.temp/<demo-stem>.<ext>`) and validates the generated XML against the XSD schemas. Single format: append `docx` / `pptx` / `xlsx`; single file: subcommands like `slide <path>`.
- **JSON Schema (AI-facing)**: the TS `*Options` types are frozen into draft-07 schemas at `packages/office-open/schemas/{docx,pptx,xlsx}.schema.json` (committed artifacts).
  - Regenerate after touching public option types: `pnpm schema:generate` (root script; `schema:check` diffs against the committed files and exits non-zero on drift — CI-enforced, so a stale schema fails the build until regenerated)
  - Validate the corpus against them: `pnpm schema:validate` — office-open demo JSONs directly plus every package demo round-tripped through parse (`scripts/schema-roundtrip-worker.ts`, run inside each package dir so tsconfig aliases resolve to source). Zero errors is the bar; a failure is either a schema gap (fix generation post-processing or the source type) or genuine options/type drift.

## Measurement Units

**Every numeric field uses the unit a human would say out loud; the library owns the OOXML scale factors.** Percent = integer percent (`50` = 50%, ×1000), angle = degrees (`45` = 45°, ×60000), font size/spacing = points (`12` = 12 pt, ×100 or ×2 half-points), duration = milliseconds. No field exposes the raw `ST_Percentage`/`ST_Angle` integer to the caller. Geometry/sizing fields take **`number`** (native unit — EMU for DrawingML, twip for Word) or a **`UniversalMeasure` string** (mm/cm/in/pt/pc/pi, plus px at 96 DPI on DrawingML). Convert with `convertToEmu` / `convertToTwip` / `convertToPt` / `convertToInch` — polymorphic, `number` passes through, round-trip lossless. `UniversalMeasure` is input-only: stringify emits the integer/unit the XSD requires, so a geometric-length field stays plain `number` only when its XSD type is integer-only (bevel, 3D extrusion depth). **Stringify and parse change together** so every scale cancels out. See [CONTRIBUTING.md](./CONTRIBUTING.md#measurement-units).

## Code Conventions

Project-specific core; full standards in [CONTRIBUTING.md](./CONTRIBUTING.md).

- **Naming**: `<part>Desc` descriptors, `<Part>Options` interfaces, `stringify*()` / `parse*()` / `patch*()` helpers. One concept, one `*Options` name across packages (`PictureOptions` in docx/pptx/xlsx). No `Model`/`Content`/`Element` prefixes (`UniversalMeasure` is XSD `ST_UniversalMeasure`, not a precedent). Public-API renames: pre-1.0 rename in place (no `@deprecated` aliases — `GroupOptions`/`ConnectorOptions`/`ShapeOptions` replaced `WpgGroupRunOptions`/`ConnectorShapeOptions`/`WpsShapeRunOptions` directly); post-1.0 use `@deprecated` aliases for transition.
- **OOXML-fidelity naming**: names mirror the OOXML element they serialize, not library legacy. A picture frame (`pic:pic`/`p:pic`/`xdr:pic`) is `picture` everywhere (`PictureOptions`, `{ picture }`, `picture-run`, `createPictureData`); `image` survives only where OOXML uses it — MIME strings (`image/png`), OPC relationship URIs (`…/relationships/image`), `a:blip`/VML bitmap data, the page background (`BackgroundImageOptions`). Internal symbols take a clean break, not `@deprecated`.
- **Properties**: naming follows **domain precedent**, not length — established short forms stay short (`id`/`idx`/`numFmt`/`fontId`/`r`/`g`/`b`/`num`/`toc`), invented abbreviations spell out (`level`/`field`/`vertical`), compound abbreviations always spell out (`lnIdx` → `lineReferenceIndex`); reference elements → `*Reference`; root-vs-derived-noun pairs (`align`/`alignment`, `rotate`/`rotation`) take the form of the type name + package mainstream (object-model lib trends to the noun), not the CSS verb root; same concept, same name everywhere. See [CONTRIBUTING.md](./CONTRIBUTING.md#property-naming) for the decision rule.
- **Enums**: string literal unions by default; `as const` objects only when values are referenced at runtime (point access, `Record` keys, `Object.values`).
- **XML generation**: template-literal concatenation — no intermediate object trees.
- Convention changes (structure / naming / cross-format) land here and in CONTRIBUTING.md **before** code — docs are the target-state spec.
