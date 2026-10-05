export type LegacyCapabilityMode = "parse-only" | "read-only";
export type LegacyCapabilityLayer = "container" | "record" | "stream" | "mapper";

export interface LegacyCapabilityOwner {
  package: string;
  module: string;
  export: string;
}

export interface LegacyCapabilityMapper {
  layer: LegacyCapabilityLayer;
  module: string;
  export: string;
}

export interface LegacyCapabilityFixture {
  key: string;
  generator: string;
}

export interface LegacyCapability {
  name: string;
  owner: LegacyCapabilityOwner;
  mapper: LegacyCapabilityMapper;
  assertion: string;
  testId: string;
  fixture: LegacyCapabilityFixture;
  supported: boolean;
}

export interface LegacyCapabilityFormat {
  id: "doc" | "xls" | "ppt" | "rtf";
  mode: LegacyCapabilityMode;
  testModule: string;
  capabilities: readonly LegacyCapability[];
}

interface CapabilityInput {
  name: string;
  assertion: string;
  testId: string;
  fixtureKey?: string;
  fixtureGenerator?: string;
}

function capabilities(
  owner: LegacyCapabilityOwner,
  mapper: LegacyCapabilityMapper,
  fixtureKey: string,
  fixtureGenerator: string,
  inputs: readonly CapabilityInput[],
): LegacyCapability[] {
  return inputs.map((input) => ({
    ...input,
    owner,
    mapper,
    fixture: {
      key: input.fixtureKey ?? fixtureKey,
      generator: input.fixtureGenerator ?? fixtureGenerator,
    },
    supported: true,
  }));
}

export const LEGACY_CAPABILITY_REGISTRY: readonly LegacyCapabilityFormat[] = [
  {
    id: "doc",
    mode: "parse-only",
    testModule: "src/parser.spec.ts",
    capabilities: capabilities(
      { package: "doc", module: "src/index.ts", export: "parseDocument" },
      {
        layer: "mapper",
        module: "src/mappers/document.ts",
        export: "parseInternal",
      },
      "synthetic-doc-cfb",
      "buildDocument",
      [
        {
          name: "Word 6.0/95 FIB parsing",
          assertion: "legacy FIB text projects without modern FIB extensions",
          testId: "extracts Word 6.0/95 text without modern FIB extensions",
        },
        {
          name: "Piece table text extraction",
          assertion: "compressed and Unicode pieces retain text and ordering",
          testId: "projects compressed and Unicode pieces with character properties",
        },
        {
          name: "Field results",
          assertion: "field result text remains while instructions are omitted",
          testId: "keeps field results without field instructions",
        },
        {
          name: "Hyperlinks",
          assertion: "HYPERLINK fields become hyperlink children",
          testId: "projects HYPERLINK field results as hyperlink children",
        },
        {
          name: "Textboxes",
          assertion: "actual textbox boundaries project and reusable spares are skipped",
          testId: "projects actual textbox boundaries and skips reusable FTXBXS records",
        },
        {
          name: "Footnotes",
          assertion: "individual footnote boundaries project to section children",
          testId: "projects individual footnote boundaries",
        },
        {
          name: "Default header stream",
          assertion: "the default header story binds to the section",
          testId: "projects the default header story slot",
        },
        {
          name: "Character formatting (CHPX)",
          assertion: "full CHPX run formatting projects onto runs",
          testId: "decodes full CHPX run formatting",
        },
        {
          name: "Paragraph formatting (PAPX)",
          assertion: "alignment, indentation, spacing, and outline level project",
          testId: "decodes PAPX alignment, indentation, spacing and outline level",
        },
        {
          name: "Section properties (SEPX)",
          assertion: "SEPX geometry maps to section page size and margins",
          testId: "maps SEPX geometry to DocumentOptions.sections",
        },
        {
          name: "Tables",
          assertion: "in-table PAPX and TAP build a row",
          testId: "builds a table row from in-table PAPX and TAP",
        },
        {
          name: "Bookmarks",
          assertion: "bookmark names and ranges project as paired children",
          testId: "projects bookmark names and ranges",
        },
        {
          name: "Annotations and revision ranges",
          assertion: "annotation references become revision ranges",
          testId: "projects annotation references as canonical revision ranges",
        },
        {
          name: "Main document fields",
          assertion: "PlcffldMom field instructions are preserved",
          testId: "projects main-document field instructions into canonical field runs",
        },
        {
          name: "Pictures (BStore and Data)",
          assertion: "BStore and Data-stream BLIP data project",
          testId: "projects Escher BStore pictures into canonical inline pictures",
          fixtureKey: "synthetic-doc-escher-pictures",
        },
        {
          name: "Embedded pictures",
          assertion: "embedded BLIP data projects without consulting Data",
          testId: "projects embedded BLIP pictures without consulting the Data stream",
          fixtureKey: "synthetic-doc-embedded-picture",
        },
        {
          name: "Lists and numbering",
          assertion: "PlfLst and PlfLfo become numbering definitions",
          testId: "converts PlfLst and PlfLfo to numbering definitions",
        },
        {
          name: "Stylesheet",
          assertion: "STSH names and based-on relationships project",
          testId: "converts STSH style names and based-on relationships",
        },
      ],
    ),
  },
  {
    id: "xls",
    mode: "parse-only",
    testModule: "src/parse.spec.ts",
    capabilities: capabilities(
      { package: "xls", module: "src/index.ts", export: "parseWorkbook" },
      {
        layer: "mapper",
        module: "src/mappers/workbook.ts",
        export: "parseWorkbook",
      },
      "synthetic-xls-biff8",
      "workbook + buildSmallContainer",
      [
        {
          name: "Workbook stream lookup",
          assertion: "Workbook/Book stream selection works",
          testId: "accepts Book as the stream name",
        },
        {
          name: "Bound sheets",
          assertion: "BOUNDSHEET visibility and ordering project",
          testId: "maps sheet visibility and sorts rows",
        },
        {
          name: "Shared strings (SST)",
          assertion: "continuation records extend the shared string table",
          testId: "reads BIFF8 shared strings across continuation records",
        },
        {
          name: "Number cells",
          assertion: "BIFF2 NUMBER cells project values",
          testId: "parses a raw BIFF2 worksheet stream",
        },
        {
          name: "RK/MulRK numeric cells",
          assertion: "RK and MulRK values project",
          testId: "parses BIFF8 cells, dimensions, and cached formula results",
        },
        {
          name: "Label and LabelSst cells",
          assertion: "label cells resolve text values",
          testId: "parses BIFF8 cells, dimensions, and cached formula results",
        },
        {
          name: "RString rich text cells",
          assertion: "rich label text and runs project",
          testId: "reads BIFF8 rich label cells",
        },
        {
          name: "Boolean and error cells",
          assertion: "boolean and error cells project",
          testId: "parses boolean and error cells",
        },
        {
          name: "Blank cells",
          assertion: "blank cells retain cell coordinates",
          testId: "parses BIFF8 cells, dimensions, and cached formula results",
        },
        {
          name: "Formula cached values",
          assertion: "formula cells retain cached results",
          testId: "parses BIFF8 cells, dimensions, and cached formula results",
        },
        {
          name: "Shared and array formulas",
          assertion: "normal, shared, and array formula tokens project",
          testId: "parses normal, shared and array formula tokens",
        },
        {
          name: "Fonts, formats, and XF styles",
          assertion: "FONT, FORMAT, and XF formatting project",
          testId: "maps FONT, FORMAT and XF to cell formatting",
        },
        {
          name: "Merged cells",
          assertion: "MERGEDCELLS ranges project",
          testId: "parses MERGEDCELLS and COLINFO",
        },
        {
          name: "Column information",
          assertion: "COLINFO widths project",
          testId: "parses MERGEDCELLS and COLINFO",
        },
        {
          name: "Row layout",
          assertion: "row heights and metadata project",
          testId: "reads row layout semantics",
        },
        {
          name: "Defined names",
          assertion: "NAME records project defined names",
          testId: "parses NAME defined names",
        },
        {
          name: "External references",
          assertion: "SUPBOOK and EXTERNSHEET project",
          testId: "parses internal and external SUPBOOK references",
        },
        {
          name: "DDE and OLE references",
          assertion: "DDE and OLE SUPBOOK targets project",
          testId: "parses DDE and OLE SUPBOOK references",
        },
        {
          name: "Hyperlinks",
          assertion: "HLINK targets project",
          testId: "parses HLINK hyperlinks",
        },
        {
          name: "Notes and comments",
          assertion: "OBJ/TXO/NOTE text projects",
          testId: "parses OBJ/TXO/NOTE comments",
        },
        {
          name: "Drawings",
          assertion: "MSODRAWING Escher anchors project",
          testId: "recursively parses MSODRAWING Escher anchors",
        },
        {
          name: "Conditional formatting",
          assertion: "CONDFMT and CF rules project",
          testId: "parses CONDFMT and CF rules",
        },
        {
          name: "Page setup",
          assertion: "headers, footers, margins, and setup project",
          testId: "parses HEADER, FOOTER, margins and PAGESETUP",
        },
        {
          name: "Codepage and date mode",
          assertion: "CODEPAGE strings and DATEMODE dates project",
          testId: "decodes CODEPAGE strings and DATEMODE",
        },
        { name: "Palette", assertion: "PALETTE colors project", testId: "parses PALETTE colors" },
        {
          name: "Malformed auxiliary formatting tolerance",
          assertion: "malformed formatting records do not lose workbook cells",
          testId: "skips malformed auxiliary formatting records",
        },
      ],
    ),
  },
  {
    id: "ppt",
    mode: "parse-only",
    testModule: "src/parser.spec.ts",
    capabilities: capabilities(
      { package: "ppt", module: "src/index.ts", export: "parsePresentation" },
      {
        layer: "mapper",
        module: "src/mappers/presentation.ts",
        export: "parsePresentation",
      },
      "synthetic-ppt-cfb",
      "buildDocument + buildCfb",
      [
        {
          name: "Persisted presentation container",
          assertion: "current user, edit, persist, document, slides, master, and text project",
          testId: "projects persisted slides, anchored text shapes, and paragraph runs",
        },
        {
          name: "Text line separators",
          assertion: "vertical tabs remain paragraph line separators",
          testId: "preserves vertical-tab line separators in one paragraph",
        },
        {
          name: "Notes slides",
          assertion: "SlideAtom references resolve to notes",
          testId: "projects notes associated by SlideAtom.notesIdRef",
        },
        {
          name: "Animations",
          assertion: "animation effect, target, trigger, and delay project",
          testId: "projects legacy shape animation effect, target, trigger, and delay",
        },
        {
          name: "Text hyperlinks",
          assertion: "external text hyperlinks project",
          testId: "projects legacy external text hyperlinks",
        },
        {
          name: "Escher pictures",
          assertion: "direct OfficeArt pictures and references project",
          testId: "projects direct OfficeArt picture records and blip references",
        },
        {
          name: "Group anchors",
          assertion: "nested groups retain 32-bit child anchors",
          testId: "reads 32-bit child anchors in nested groups",
        },
        {
          name: "No-text autoshapes",
          assertion: "no-text autoshapes project as base shapes with geometry",
          testId: "keeps a no-text autoshape as a base shape",
        },
        {
          name: "Legacy lines and connectors",
          assertion: "lines and connectors project as endpoint children",
          testId: "keeps a legacy line as a line child",
        },
        {
          name: "Mixed group children",
          assertion: "shape, picture, and connector order is retained in groups",
          testId: "preserves shape, picture, and connector order in groups",
        },
        {
          name: "Unknown visual properties",
          assertion: "unknown visual OPT properties fail structurally",
          testId: "reports unknown visual shape properties structurally",
        },
        {
          name: "Embedded picture stream",
          assertion: "embedded BSE pictures project and trim names",
          testId: "projects an embedded BSE picture and trims its name",
        },
        {
          name: "Picture fills and text",
          assertion: "picture-filled shapes retain text",
          testId: "keeps text on a shape with an embedded picture fill",
        },
        {
          name: "Picture corruption isolation",
          assertion: "a corrupt picture does not discard valid pictures",
          testId: "isolates a corrupt picture beside a valid embedded picture",
        },
        {
          name: "Grouped text tables",
          assertion: "grouped text grids project as tables",
          testId: "projects a complete grouped-text grid as a table",
        },
        {
          name: "Table paragraph preservation",
          assertion: "grouped table cells preserve paragraphs",
          testId: "preserves multiple paragraphs in grouped table cells",
        },
        {
          name: "Top-level text tables",
          assertion: "compact top-level grids project as tables",
          testId: "projects a compact top-level text grid as a table",
        },
        {
          name: "Table wrapper unwrapping",
          assertion: "table wrapper groups unwrap",
          testId: "unwraps a top-level table group wrapper",
        },
        {
          name: "Merged table spans",
          assertion: "merged text rectangles project as spans",
          testId: "projects merged grouped text rectangles as table spans",
        },
        {
          name: "Vertical-tab line breaks",
          assertion: "vertical tabs map without splitting paragraphs",
          testId: "maps vertical tabs to line breaks without splitting paragraphs",
        },
        {
          name: "Line-feed line breaks",
          assertion: "line feeds map without splitting paragraphs",
          testId: "maps line feeds to line breaks without splitting paragraphs",
        },
        {
          name: "Unreferenced slide text",
          assertion: "slide-list text is retained",
          testId: "projects unreferenced slide-list text",
        },
        {
          name: "Date and slide fields",
          assertion: "date and slide-number placeholders resolve",
          testId: "resolves date and slide-number field placeholders",
        },
        {
          name: "Non-CFB rejection",
          assertion: "non-CFB input is rejected",
          testId: "rejects input that is not CFB",
        },
        {
          name: "Encrypted atom rejection",
          assertion: "encrypted Current User atoms are rejected",
          testId: "rejects encrypted Current User atoms",
        },
      ],
    ),
  },
  {
    id: "rtf",
    mode: "read-only",
    testModule: "src/parser.spec.ts",
    capabilities: capabilities(
      { package: "rtf", module: "src/index.ts", export: "parseDocument" },
      {
        layer: "mapper",
        module: "src/mappers/document.ts",
        export: "parseRtfTokens",
      },
      "synthetic-rtf-source",
      "inline RTF string fixtures",
      [
        {
          name: "Rich text extraction",
          assertion: "standard text becomes document paragraphs",
          testId: "parses standard text into a document paragraph",
        },
        {
          name: "Character formatting",
          assertion: "bold and paragraph alignment project",
          testId: "scopes bold formatting to its group and maps paragraph alignment",
        },
        {
          name: "Fonts and colors",
          assertion: "fonts, sizes, colors, and underline project",
          testId: "maps fonts, sizes, colors, and underline",
        },
        {
          name: "Unicode and hexadecimal escapes",
          assertion: "tabs, escapes, and controls decode",
          testId: "parses tabs, hexadecimal, Unicode, and common character controls",
        },
        {
          name: "Tables",
          assertion: "cell boundaries and column widths project",
          testId: "projects cell boundaries and twip column widths",
        },
        {
          name: "Headers, footers, and footnotes",
          assertion: "section stories project",
          testId: "projects headers, footers, and footnotes",
        },
        {
          name: "Fields",
          assertion: "HYPERLINK, PAGE, and DATE fields project",
          testId: "projects HYPERLINK, PAGE, and DATE fields",
        },
        {
          name: "Pictures",
          assertion: "PNG bytes and dimensions project",
          testId: "extracts PNG picture bytes and twip dimensions",
        },
        {
          name: "Binary destinations",
          assertion: "picture and object binary payloads project",
          testId: "recognizes an embedded object class",
        },
        {
          name: "Stylesheets",
          assertion: "complete style metadata and formatting project",
          testId: "projects stylesheet metadata and body formatting",
        },
        {
          name: "Numbering tables",
          assertion: "complete list tables and overrides project",
          testId: "projects a complete list table and override",
        },
        {
          name: "Bookmarks",
          assertion: "bookmark starts and ends pair",
          testId: "pairs bookmark starts and ends",
        },
        {
          name: "Lists",
          assertion: "list references and levels project",
          testId: "projects list references and levels",
        },
      ],
    ),
  },
] as const;
