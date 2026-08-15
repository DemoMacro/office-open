/**
 * Entry catalog for schema slicing and lookup.
 *
 * The catalog is the single source of truth for two concerns:
 *   1. the recommended lookup index (`office-open schema index`)
 *   2. the stub boundary during closure traversal — any catalog member that
 *      is NOT requested collapses to an expandable stub instead of being
 *      walked, which keeps slices small (a full DocumentOptions closure
 *      covers ~440 of 442 definitions).
 *
 * An entry is a domain root a user would construct independently: big enough
 * that stubbing it saves real tokens. Small helper types (BookmarkOptions,
 * SimpleFieldOptions, …) are deliberately NOT cataloged — they always expand
 * inside any slice, so they never need a second lookup.
 *
 * @module
 */

import type { DocumentType } from "./schemas";

/** One cataloged lookup entry. */
export interface SchemaEntry {
  /** Definition name in the format's schema `definitions` (the TS type name). */
  name: string;
  /** Coarse domain grouping used by `office-open schema index` output. */
  domain: string;
  /** One-line summary (hand-written; schema descriptions may be missing). */
  summary: string;
}

/** Root definition of each format's options tree. */
export const SCHEMA_ROOTS: Record<DocumentType, string> = {
  docx: "DocumentOptions",
  pptx: "PresentationOptions",
  xlsx: "WorkbookOptions",
};

const DOCX_ENTRIES: readonly SchemaEntry[] = [
  {
    name: "DocumentOptions",
    domain: "document",
    summary: "Root document options (sections, styles, numbering, settings)",
  },
  {
    name: "AppPropertiesOptions",
    domain: "document",
    summary: "Core + app metadata (title, creator, dates)",
  },
  { name: "SettingsOptions", domain: "document", summary: "w:settings document behavior flags" },
  {
    name: "FeaturesOptions",
    domain: "document",
    summary: "Track changes, comments visibility, update fields",
  },
  {
    name: "CompatibilityOptions",
    domain: "document",
    summary: "Legacy Word compatibility settings",
  },
  {
    name: "DocumentBackgroundOptions",
    domain: "document",
    summary: "Page background (color or image)",
  },
  {
    name: "MailMergeOptions",
    domain: "document",
    summary: "Mail merge envelope (dataSource, fields)",
  },
  { name: "BibliographyOptions", domain: "document", summary: "Bibliography + citation sources" },
  {
    name: "SectionOptions",
    domain: "structure",
    summary: "One section: properties, headers/footers, children",
  },
  {
    name: "SectionPropertiesOptions",
    domain: "structure",
    summary: "Page size, margins, columns, grid",
  },
  {
    name: "HeaderFooterGroup<HeaderFooterReference>",
    domain: "structure",
    summary: "Even/first/default header+footer set (quote the name in shells)",
  },
  {
    name: "ParagraphOptions",
    domain: "text",
    summary: "Block paragraph (heading, alignment, spacing, runs)",
  },
  { name: "RunOptions", domain: "text", summary: "Text run (font, color, size, emphasis)" },
  {
    name: "ParagraphRunOptions",
    domain: "text",
    summary: "Run options including child-run wrappers (breaks, tabs, fields)",
  },
  { name: "HyperlinkOptions", domain: "text", summary: "Hyperlink with child runs" },
  { name: "MathInput", domain: "text", summary: "OMML math equation tree" },
  { name: "TableOptions", domain: "tables", summary: "Table (rows, width, borders, layout)" },
  {
    name: "TableRowOptions",
    domain: "tables",
    summary: "Table row (cells, height, header repeat)",
  },
  {
    name: "TableCellOptions",
    domain: "tables",
    summary: "Table cell (span, shading, vertical align)",
  },
  {
    name: "TableStyleOptions",
    domain: "tables",
    summary: "Table style banding + conditional formats",
  },
  {
    name: "PictureOptions",
    domain: "media",
    summary: "Inline/anchored image (data, size, altText)",
  },
  { name: "ChartOptions", domain: "media", summary: "Chart (type, series, axes, legend)" },
  {
    name: "ChartSpaceOptions",
    domain: "media",
    summary: "Full chartSpace envelope wrapping chart options",
  },
  {
    name: "SmartArtOptions",
    domain: "media",
    summary: "Diagram nodes + layout/style/color references",
  },
  {
    name: "GroupOptions",
    domain: "media",
    summary: "Drawing group container (children, transforms)",
  },
  {
    name: "ShapeOptions",
    domain: "media",
    summary: "Floating text box / shape with drawing anchor",
  },
  {
    name: "SdtRunOptions",
    domain: "media",
    summary: "Structured document tag wrapping run content",
  },
  { name: "StylesOptions", domain: "styles", summary: "Default + paragraph/character/link styles" },
  {
    name: "StyleDefinitionOptions",
    domain: "styles",
    summary: "Base style fields shared by style kinds",
  },
  { name: "ParagraphStyleOptions", domain: "styles", summary: "Paragraph style definition" },
  { name: "CharacterStyleOptions", domain: "styles", summary: "Character (run) style definition" },
  {
    name: "NumberingOptions",
    domain: "styles",
    summary: "Bullet/numbering definitions and references",
  },
  { name: "CommentsOptions", domain: "annotations", summary: "Comments collection" },
  {
    name: "CommentOptions",
    domain: "annotations",
    summary: "One comment (author, children, date)",
  },
  {
    name: "FootnotePropertiesOptions",
    domain: "annotations",
    summary: "Footnotes section options",
  },
  { name: "EndnotePropertiesOptions", domain: "annotations", summary: "Endnotes section options" },
];

const PPTX_ENTRIES: readonly SchemaEntry[] = [
  {
    name: "PresentationOptions",
    domain: "presentation",
    summary: "Root presentation options (slides, size, masters)",
  },
  { name: "AppPropertiesOptions", domain: "presentation", summary: "Core + app metadata" },
  { name: "ShowOptions", domain: "presentation", summary: "Slide show settings" },
  {
    name: "CustomShowOptions",
    domain: "presentation",
    summary: "Custom slide show (name, slide list)",
  },
  {
    name: "PrintPropertiesOptions",
    domain: "presentation",
    summary: "Handouts/notes printing settings",
  },
  { name: "SlideOptions", domain: "slides", summary: "One slide (children, background, notes)" },
  { name: "BackgroundOptions", domain: "slides", summary: "Slide background fill" },
  {
    name: "SlideHeaderFooterOptions",
    domain: "slides",
    summary: "Slide header/footer/slide-number placeholders",
  },
  {
    name: "TransitionOptions",
    domain: "slides",
    summary: "Slide transition (type, duration, triggers)",
  },
  { name: "SlideAnimation", domain: "slides", summary: "Slide animation timeline root" },
  {
    name: "AnimationOptions",
    domain: "slides",
    summary: "Animation effect (target, effect, timing)",
  },
  {
    name: "MasterDefinition",
    domain: "masters",
    summary: "Slide master definition (placeholders, styles)",
  },
  {
    name: "ThemeOptions",
    domain: "masters",
    summary: "Theme (color scheme, fonts, format scheme)",
  },
  { name: "NotesMasterOptions", domain: "masters", summary: "Notes master options" },
  { name: "NotesSlideOptions", domain: "masters", summary: "Notes slide content" },
  { name: "HandoutMasterOptions", domain: "masters", summary: "Handout master options" },
  {
    name: "ShapeOptions",
    domain: "shapes",
    summary: "Auto shape (geometry, fill, outline, textBody)",
  },
  { name: "LineShapeOptions", domain: "shapes", summary: "Straight line shape" },
  { name: "ConnectorOptions", domain: "shapes", summary: "Connector with optional endpoints" },
  { name: "GroupOptions", domain: "shapes", summary: "Group container (children, transforms)" },
  {
    name: "TextBodyOptions",
    domain: "shapes",
    summary: "Shape text body (paragraphs, list styles)",
  },
  { name: "BulletOptions", domain: "shapes", summary: "Paragraph bullet character/numbering" },
  { name: "TextListStyleOptions", domain: "shapes", summary: "List style per outline level" },
  { name: "PictureOptions", domain: "media", summary: "Image (data, size, style, crop)" },
  { name: "TableOptions", domain: "media", summary: "Table (rows, first-row/banding flags)" },
  { name: "ChartOptions", domain: "media", summary: "Chart (type, series, axes, legend)" },
  {
    name: "SmartArtOptions",
    domain: "media",
    summary: "Diagram nodes + layout/style/color references",
  },
  { name: "VideoFrameOptions", domain: "media", summary: "Video frame (data, poster, trim)" },
  { name: "AudioFrameOptions", domain: "media", summary: "Audio frame (data, icon, playback)" },
  { name: "OleOptions", domain: "media", summary: "OLE object embedding" },
];

const XLSX_ENTRIES: readonly SchemaEntry[] = [
  {
    name: "WorkbookOptions",
    domain: "workbook",
    summary: "Root workbook options (worksheets, properties)",
  },
  { name: "AppPropertiesOptions", domain: "workbook", summary: "Core + app metadata" },
  {
    name: "WorkbookPropertiesOptions",
    domain: "workbook",
    summary: "Workbook flags (date1904, iterateCalc)",
  },
  {
    name: "WorkbookProtectionOptions",
    domain: "workbook",
    summary: "Workbook structure/window protection",
  },
  {
    name: "CalculationPropertiesOptions",
    domain: "workbook",
    summary: "Calculation mode and iteration",
  },
  { name: "DefinedNameOptions", domain: "workbook", summary: "Named range/formula" },
  {
    name: "WorksheetOptions",
    domain: "sheets",
    summary: "One worksheet (rows, columns, merges, views)",
  },
  { name: "ChartsheetOptions", domain: "sheets", summary: "Chart sheet hosting one chart" },
  { name: "SheetProtectionOptions", domain: "sheets", summary: "Sheet protection flags" },
  { name: "SheetViewOptions", domain: "sheets", summary: "Sheet view (zoom, selection, pane)" },
  {
    name: "HeaderFooterOptions",
    domain: "sheets",
    summary: "Print header/footer (odd/even/first)",
  },
  {
    name: "PageSetupOptions",
    domain: "sheets",
    summary: "Print page setup (orientation, scale, paper)",
  },
  { name: "PageMarginsOptions", domain: "sheets", summary: "Print page margins" },
  { name: "PrintOptions", domain: "sheets", summary: "Print gridlines/headings/order" },
  { name: "RowOptions", domain: "data", summary: "Row (cells, height, style)" },
  { name: "CellOptions", domain: "data", summary: "Cell (value, style, formula)" },
  { name: "RichTextOptions", domain: "data", summary: "Cell rich text runs" },
  { name: "CommentOptions", domain: "data", summary: "Cell comment" },
  { name: "HyperlinkOptions", domain: "data", summary: "Cell hyperlink" },
  { name: "ColumnOptions", domain: "data", summary: "Column (width, style, custom)" },
  {
    name: "StyleOptions",
    domain: "style",
    summary: "Cell style (font, fill, border, alignment, numFmt)",
  },
  { name: "FontOptions", domain: "style", summary: "Font (name, size, color, emphasis)" },
  { name: "FillOptions", domain: "style", summary: "Pattern or gradient fill" },
  { name: "BorderOptions", domain: "style", summary: "Border sides and diagonals" },
  { name: "AlignmentOptions", domain: "style", summary: "Text alignment and wrap" },
  { name: "AutoFilterOptions", domain: "features", summary: "Auto filter range and columns" },
  { name: "DataValidationOptions", domain: "features", summary: "Cell input validation rules" },
  { name: "ConditionalFormatOptions", domain: "features", summary: "Conditional formatting rules" },
  { name: "TableOptions", domain: "features", summary: "Worksheet table (range, columns, style)" },
  { name: "PivotTableOptions", domain: "features", summary: "Pivot table definition" },
  { name: "WorksheetChartOptions", domain: "features", summary: "Anchored chart on a worksheet" },
  { name: "PictureOptions", domain: "media", summary: "Anchored image" },
  { name: "ShapeOptions", domain: "media", summary: "Anchored shape" },
  { name: "ConnectorOptions", domain: "media", summary: "Anchored connector" },
  { name: "GroupOptions", domain: "media", summary: "Anchored group container" },
  { name: "ThemeOptions", domain: "media", summary: "Workbook theme" },
];

/** Cataloged lookup entries per format. */
export const SCHEMA_ENTRIES: Readonly<Record<DocumentType, readonly SchemaEntry[]>> = {
  docx: DOCX_ENTRIES,
  pptx: PPTX_ENTRIES,
  xlsx: XLSX_ENTRIES,
};

/** Definition names that act as stub boundaries when not requested. */
export function entryNames(type: DocumentType): ReadonlySet<string> {
  return new Set(SCHEMA_ENTRIES[type].map((entry) => entry.name));
}
