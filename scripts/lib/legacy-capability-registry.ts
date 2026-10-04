export type LegacyCapabilityMode = "parse-only" | "read-only";

export interface LegacyCapabilityOwner {
  package: string;
  module: string;
  export: string;
}

export interface LegacyCapability {
  name: string;
  owner: LegacyCapabilityOwner;
  supported: boolean;
}

export interface LegacyCapabilityFormat {
  id: "doc" | "xls" | "ppt" | "rtf";
  mode: LegacyCapabilityMode;
  capabilities: readonly LegacyCapability[];
}

function capability(
  format: LegacyCapabilityFormat["id"],
  exportName: string,
  names: readonly string[],
): LegacyCapability[] {
  const owner = { package: format, module: "src/index.ts", export: exportName };
  return names.map((name) => ({ name, owner, supported: true }));
}

export const LEGACY_CAPABILITY_REGISTRY: readonly LegacyCapabilityFormat[] = [
  {
    id: "doc",
    mode: "parse-only",
    capabilities: capability("doc", "parseDocument", [
      "FIB version/nFib parsing",
      "Piece table text extraction",
      "Unicode/ANSI piece mapping",
      "Encryption detection",
      "Legacy codepage mapping",
      "Bookmarks",
      "Comments/annotations",
      "Tracked changes",
      "Fields",
      "Hyperlinks",
      "Footnotes/endnotes",
      "Lists/numbering",
      "Styles (stylesheet)",
      "Character formatting (CHPX run properties)",
      "Paragraph formatting (PAPX)",
      "Section properties (SEPX)",
      "Tables",
      "Headers/footers streams",
      "Pictures (Escher/Data)",
    ]),
  },
  {
    id: "xls",
    mode: "parse-only",
    capabilities: capability("xls", "parseWorkbook", [
      "Workbook globals/BOF",
      "Bound sheets (BOUNDSHEET)",
      "Shared strings (SST)",
      "Number cells (NUMBER)",
      "RK/MulRK numeric cells",
      "Label/LabelSst text cells",
      "RString cells",
      "Bool/error cells",
      "Blank cells",
      "Formula cells + cached values",
      "Shared formulas (SHRFMLA)",
      "Array formulas (ARRAY)",
      "XF style records",
      "Number formats (FORMAT)",
      "Fonts (FONT)",
      "Merged cells (MERGEDCELLS)",
      "Column width/info (COLINFO)",
      "Row height (ROW)",
      "Defined names (NAME)",
      "External references (SUPBOOK/EXTERNSHEET)",
      "Hyperlinks (HLINK)",
      "Notes/comments (NOTE/TXO)",
      "Drawings (MSODRAWING)",
      "Conditional formatting (CONDFMT/CF)",
      "Page setup",
      "Codepage/encoding",
      "Date mode (1904)",
      "Palette (PALETTE)",
    ]),
  },
  {
    id: "ppt",
    mode: "parse-only",
    capabilities: capability("ppt", "parsePresentation", [
      "Current user atom",
      "User edit atom",
      "Persist pointer blocks",
      "Document container",
      "Slides (SlideContainer)",
      "Slide atoms",
      "Main masters",
      "Text extraction (TextChars/TextBytes)",
      "Text header/formatting",
      "Slide list with text",
      "Escher shapes",
      "Escher shape properties",
      "Escher pictures (BSE/BStore)",
      "Escher groups/anchors",
      "Textboxes (ClientTextbox)",
      "Drawing containers (PPDrawing)",
      "Date/time atoms",
      "Slide number atoms",
      "Notes slides",
      "Animations",
      "Encryption (DocumentEncryptionAtom)",
      "Outline text references",
    ]),
  },
  {
    id: "rtf",
    mode: "read-only",
    capabilities: capability("rtf", "parseRtf", [
      "Rich text extraction",
      "Character formatting",
      "Fonts and colors",
      "Unicode and hexadecimal escapes",
    ]),
  },
] as const;
