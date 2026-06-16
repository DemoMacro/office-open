/**
 * Declarative part registry — the machine-readable source of truth for which
 * parts each OOXML package is expected to contain.
 *
 * Consumed by {@link validateOpcConsistency} to detect two failure classes the
 * XSD single-part validator cannot see:
 *   - O1 orphan part  — a part present in the ZIP that no rule accounts for
 *   - O2 missing part — an `always` part the ZIP lacks
 *
 * Content-type and relationship integrity (O3/O5/O6/O7) are derived directly
 * from the ZIP + `.rels` + `[Content_Types].xml`, so they need no registry.
 *
 * Presence semantics:
 *   - `always`      — present in *any* valid package (content-types map,
 *                     root rels, main part only). O2 fires when absent. Kept
 *                     minimal so template/round-trip packages — whose
 *                     `[Content_Types]` and part set are passed through from a
 *                     source rather than rebuilt — do not false-positive.
 *   - `conditional` — emitted by a fresh `compileDocument` run but a
 *                     round-tripped/template package may omit it. Absence is
 *                     legitimate; O6 (stale Override) still catches the case
 *                     where its content type is declared but the part is gone.
 *   - `repeated`    — one per index (slides, worksheets, headers, …).
 *
 * Reference: ECMA-376 Part 2 (OPC). Content-type values mirror the per-package
 * `content-types.ts` builders; relationship `@Type` URLs mirror
 * {@link RelationshipType}.
 *
 * @module
 */

export type PartPresence =
  | { readonly kind: "always" }
  | { readonly kind: "conditional"; readonly flag: string }
  | { readonly kind: "repeated"; readonly countFrom: string };

export interface PartDef {
  /**
   * ZIP path template. `${i}` expands per repeated index (1-based); a template
   * without the placeholder denotes a singleton part.
   */
  readonly path: string;
  /**
   * `[Content_Types].xml` Override value the part carries when generated.
   * `undefined` when the part relies on a `<Default>` extension mapping
   * (e.g. media, fonts, docx theme under raw-part passthrough).
   */
  readonly contentType?: string;
  readonly presence: PartPresence;
}

export interface PackagePartRegistry {
  readonly format: "docx" | "pptx" | "xlsx";
  readonly parts: readonly PartDef[];
  /**
   * Path prefixes that are always legitimate even when undeclared — media,
   * fonts, embeddings, altChunks, custom XML, and `.rels` parts. Used to
   * suppress O1 false positives from round-tripped / pass-through content.
   */
  readonly orphanWhitelist: readonly string[];
}

// ── DOCX ────────────────────────────────────────────────────────────────────

export const DOCX_PARTS = {
  format: "docx",
  orphanWhitelist: [
    "word/media/",
    "word/fonts/",
    "word/embeddings/",
    "word/afchunks/",
    "customXml/",
    "_rels/",
    "word/_rels/",
    "docProps/",
    "[Content_Types].xml",
  ],
  parts: [
    { path: "[Content_Types].xml", presence: { kind: "always" } },
    { path: "_rels/.rels", presence: { kind: "always" } },
    {
      path: "word/document.xml",
      contentType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
      presence: { kind: "always" },
    },
    {
      path: "word/styles.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
      presence: { kind: "conditional", flag: "fresh compile" },
    },
    {
      path: "word/numbering.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
      presence: { kind: "conditional", flag: "fresh compile" },
    },
    {
      path: "word/footnotes.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
      presence: { kind: "conditional", flag: "fresh compile" },
    },
    {
      path: "word/endnotes.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
      presence: { kind: "conditional", flag: "fresh compile" },
    },
    {
      path: "word/settings.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
      presence: { kind: "conditional", flag: "fresh compile" },
    },
    {
      path: "word/fontTable.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
      presence: { kind: "conditional", flag: "fresh compile" },
    },
    {
      path: "docProps/core.xml",
      contentType: "application/vnd.openxmlformats-package.core-properties+xml",
      presence: { kind: "conditional", flag: "fresh compile" },
    },
    {
      path: "docProps/app.xml",
      contentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
      presence: { kind: "conditional", flag: "fresh compile" },
    },
    {
      path: "docProps/custom.xml",
      contentType: "application/vnd.openxmlformats-officedocument.custom-properties+xml",
      presence: { kind: "conditional", flag: "customProperties" },
    },
    {
      path: "word/comments.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
      presence: { kind: "conditional", flag: "comments.children.length > 0" },
    },
    {
      path: "word/header${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
      presence: { kind: "repeated", countFrom: "headerCount" },
    },
    {
      path: "word/footer${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
      presence: { kind: "repeated", countFrom: "footerCount" },
    },
    {
      path: "word/charts/chart${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
      presence: { kind: "repeated", countFrom: "chartCount" },
    },
    {
      path: "word/diagrams/data${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
      presence: { kind: "repeated", countFrom: "smartArtCount" },
    },
    {
      path: "word/diagrams/layout${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
      presence: { kind: "repeated", countFrom: "smartArtCount" },
    },
    {
      path: "word/diagrams/quickStyle${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
      presence: { kind: "repeated", countFrom: "smartArtCount" },
    },
    {
      path: "word/diagrams/colors${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
      presence: { kind: "repeated", countFrom: "smartArtCount" },
    },
    {
      path: "word/diagrams/drawing${i}.xml",
      contentType: "application/vnd.ms-office.drawingml.diagramDrawing+xml",
      presence: { kind: "repeated", countFrom: "smartArtCount" },
    },
    {
      path: "word/bibliography.xml",
      contentType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.bibliography+xml",
      presence: { kind: "conditional", flag: "bibliography" },
    },
    {
      path: "word/glossary/document.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.glossary+xml",
      presence: { kind: "conditional", flag: "glossary" },
    },
    {
      path: "word/webSettings.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
      presence: { kind: "conditional", flag: "webSettings" },
    },
    // docx theme is raw-part passthrough — no fixed Override; relies on the
    // `xml` Default. contentType intentionally omitted.
    { path: "word/theme/theme1.xml", presence: { kind: "conditional", flag: "rawParts theme" } },
  ],
} as const satisfies PackagePartRegistry;

// ── PPTX ────────────────────────────────────────────────────────────────────

export const PPTX_PARTS = {
  format: "pptx",
  orphanWhitelist: [
    "ppt/media/",
    "ppt/embeddings/",
    "_rels/",
    "ppt/_rels/",
    "ppt/slideMasters/_rels/",
    "ppt/slideLayouts/_rels/",
    "ppt/slides/_rels/",
    "ppt/notesMasters/_rels/",
    "ppt/notesSlides/_rels/",
    "ppt/charts/_rels/",
    "ppt/diagrams/_rels/",
    "docProps/",
    "[Content_Types].xml",
  ],
  parts: [
    { path: "[Content_Types].xml", presence: { kind: "always" } },
    { path: "_rels/.rels", presence: { kind: "always" } },
    {
      path: "ppt/presentation.xml",
      contentType:
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
      presence: { kind: "always" },
    },
    {
      path: "docProps/core.xml",
      contentType: "application/vnd.openxmlformats-package.core-properties+xml",
      presence: { kind: "conditional", flag: "fresh compile" },
    },
    {
      path: "docProps/app.xml",
      contentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
      presence: { kind: "conditional", flag: "fresh compile" },
    },
    {
      // theme1 backs the slide master (static); theme2/theme3 back the
      // notes/handout masters when present — all share the theme+xml type.
      path: "ppt/theme/theme${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.theme+xml",
      presence: { kind: "repeated", countFrom: "masters + notes/handout masters" },
    },
    {
      path: "ppt/presProps.xml",
      contentType: "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml",
      presence: { kind: "conditional", flag: "fresh compile" },
    },
    {
      path: "ppt/viewProps.xml",
      contentType: "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml",
      presence: { kind: "conditional", flag: "fresh compile" },
    },
    {
      path: "ppt/tableStyles.xml",
      contentType: "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml",
      presence: { kind: "conditional", flag: "fresh compile" },
    },
    {
      path: "ppt/slideMasters/slideMaster${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
      presence: { kind: "repeated", countFrom: "masters.length" },
    },
    {
      path: "ppt/slideLayouts/slideLayout${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
      presence: { kind: "repeated", countFrom: "layouts.length" },
    },
    {
      path: "ppt/slides/slide${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
      presence: { kind: "repeated", countFrom: "slides.length" },
    },
    {
      path: "ppt/notesMasters/notesMaster1.xml",
      contentType: "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml",
      presence: { kind: "conditional", flag: "any slide has notes" },
    },
    {
      path: "ppt/handoutMasters/handoutMaster1.xml",
      contentType: "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml",
      presence: { kind: "conditional", flag: "includeHandoutMaster" },
    },
    {
      path: "ppt/notesSlides/notesSlide${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
      presence: { kind: "repeated", countFrom: "slides with notes" },
    },
    {
      path: "ppt/commentAuthors.xml",
      contentType:
        "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml",
      presence: { kind: "conditional", flag: "any slide has comments" },
    },
    {
      path: "ppt/comments/comment${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.presentationml.comments+xml",
      presence: { kind: "repeated", countFrom: "slides with comments" },
    },
    {
      path: "ppt/charts/chart${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
      presence: { kind: "repeated", countFrom: "charts" },
    },
    {
      path: "ppt/diagrams/data${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
      presence: { kind: "repeated", countFrom: "smartArts" },
    },
    {
      path: "ppt/diagrams/layout${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
      presence: { kind: "repeated", countFrom: "smartArts" },
    },
    {
      path: "ppt/diagrams/quickStyle${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
      presence: { kind: "repeated", countFrom: "smartArts" },
    },
    {
      path: "ppt/diagrams/colors${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
      presence: { kind: "repeated", countFrom: "smartArts" },
    },
    {
      path: "ppt/diagrams/drawing${i}.xml",
      contentType: "application/vnd.ms-office.drawingml.diagramDrawing+xml",
      presence: { kind: "repeated", countFrom: "smartArts" },
    },
    {
      path: "ppt/slideSyncPr/slideSyncPr${i}.xml",
      contentType:
        "application/vnd.openxmlformats-officedocument.presentationml.slideSyncProperties+xml",
      presence: { kind: "repeated", countFrom: "slides with slideSync" },
    },
  ],
} as const satisfies PackagePartRegistry;

// ── XLSX ────────────────────────────────────────────────────────────────────

export const XLSX_PARTS = {
  format: "xlsx",
  orphanWhitelist: [
    "xl/media/",
    "xl/embeddings/",
    "_rels/",
    "xl/_rels/",
    "xl/worksheets/_rels/",
    "xl/chartsheets/_rels/",
    "xl/drawings/_rels/",
    "xl/pivotTables/_rels/",
    "xl/pivotCache/_rels/",
    "xl/externalLinks/_rels/",
    "docProps/",
    "[Content_Types].xml",
  ],
  parts: [
    { path: "[Content_Types].xml", presence: { kind: "always" } },
    { path: "_rels/.rels", presence: { kind: "always" } },
    {
      path: "xl/workbook.xml",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
      presence: { kind: "always" },
    },
    {
      path: "docProps/core.xml",
      contentType: "application/vnd.openxmlformats-package.core-properties+xml",
      presence: { kind: "conditional", flag: "fresh compile" },
    },
    {
      path: "docProps/app.xml",
      contentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
      presence: { kind: "conditional", flag: "fresh compile" },
    },
    {
      path: "xl/styles.xml",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
      presence: { kind: "conditional", flag: "fresh compile" },
    },
    {
      path: "xl/theme/theme1.xml",
      contentType: "application/vnd.openxmlformats-officedocument.theme+xml",
      presence: { kind: "conditional", flag: "fresh compile" },
    },
    {
      // Override is emitted unconditionally; the part file is only written
      // when the shared-strings table is non-empty — an O6 stale-Override
      // risk this registry exists to surface.
      path: "xl/sharedStrings.xml",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
      presence: { kind: "conditional", flag: "sharedStrings.count > 0" },
    },
    {
      path: "xl/worksheets/sheet${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
      presence: { kind: "repeated", countFrom: "worksheets.length" },
    },
    {
      path: "xl/chartsheets/sheet${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
      presence: { kind: "repeated", countFrom: "chartsheets.length" },
    },
    {
      path: "xl/drawings/drawing${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.drawing+xml",
      presence: { kind: "conditional", flag: "worksheet has drawing" },
    },
    {
      path: "xl/comments${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
      presence: { kind: "conditional", flag: "worksheet.comments.length > 0" },
    },
    {
      // Legacy VML drawing backing older comment anchors; relies on the `vml`
      // Default added by ContentTypes.addVmlDrawing().
      path: "xl/drawings/vmlDrawing${i}.vml",
      presence: { kind: "conditional", flag: "worksheet.comments (legacy VML)" },
    },
    {
      path: "xl/charts/chart${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
      presence: { kind: "repeated", countFrom: "charts" },
    },
    {
      path: "xl/pivotTables/pivotTable${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
      presence: { kind: "repeated", countFrom: "pivotTables" },
    },
    {
      path: "xl/pivotCache/pivotCacheDefinition${i}.xml",
      contentType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
      presence: { kind: "repeated", countFrom: "pivotCaches" },
    },
    {
      path: "xl/pivotCache/pivotCacheRecords${i}.xml",
      contentType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml",
      presence: { kind: "repeated", countFrom: "pivotCaches" },
    },
    {
      path: "xl/tables/table${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
      presence: { kind: "repeated", countFrom: "tables" },
    },
    {
      path: "xl/externalLinks/externalLink${i}.xml",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml",
      presence: { kind: "repeated", countFrom: "externalLinks.length" },
    },
    {
      path: "xl/calcChain.xml",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
      presence: { kind: "conditional", flag: "any formula cell" },
    },
  ],
} as const satisfies PackagePartRegistry;

export const PART_REGISTRIES: Record<PackagePartRegistry["format"], PackagePartRegistry> = {
  docx: DOCX_PARTS,
  pptx: PPTX_PARTS,
  xlsx: XLSX_PARTS,
};
