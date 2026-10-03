/**
 * Legacy Binary Format Coverage Analysis Tool
 *
 * Unlike OOXML/ODF, the legacy binary formats (DOC, XLS, PPT) do not expose a
 * single element enumeration — DOC organizes data through FIB streams and
 * format pages rather than records, so a flat record-type list would mislead.
 * This tool measures coverage against the semantic feature set defined by the
 * MS-DOC / MS-XLS / MS-PPT structure specifications (the same surface mature
 * parsers such as xlrd, calamine, and Apache POI implement).
 *
 * A feature counts as implemented when any of its detection patterns
 * (record constants, structure names) appear in the package sources.
 *
 * Usage:
 *   pnpm tsx scripts/legacy-coverage.ts           # full report
 *   pnpm tsx scripts/legacy-coverage.ts xls       # one format only
 *   pnpm tsx scripts/legacy-coverage.ts --missing # show missing features (default)
 *   pnpm tsx scripts/legacy-coverage.ts --summary # only summary table
 */
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT_DIR = path.resolve(__dirname, "..");

interface Feature {
  name: string;
  patterns: string[];
}

const DOC_FEATURES: Feature[] = [
  { name: "FIB version/nFib parsing", patterns: ["nFib", "FIB_SIGNATURE"] },
  { name: "Piece table text extraction", patterns: ["piece", "plcfpcd"] },
  { name: "Unicode/ANSI piece mapping", patterns: ["UNICODE_PIECE"] },
  { name: "Encryption detection", patterns: ["FLAG_ENCRYPTED", "FILEPASS", "fibEncrypted"] },
  { name: "Legacy codepage mapping", patterns: ["cp1252", "20ac"] },
  { name: "Bookmarks", patterns: ["bookmark"] },
  { name: "Comments/annotations", patterns: ["plcfandRef", "annotationRef"] },
  { name: "Tracked changes", patterns: ["revision"] },
  { name: "Fields", patterns: ["plcffld"] },
  { name: "Hyperlinks", patterns: ["plcfhdd", "hyp"] },
  { name: "Footnotes/endnotes", patterns: ["footnote", "endnote"] },
  { name: "Lists/numbering", patterns: ["plcflfo", "lvlList"] },
  { name: "Styles (stylesheet)", patterns: ["stshf", "stylename"] },
  { name: "Character formatting (CHPX run properties)", patterns: ["sprmCP", "chpxText"] },
  { name: "Paragraph formatting (PAPX)", patterns: ["sprmPF", "papxText"] },
  { name: "Section properties (SEPX)", patterns: ["plcfsepx", "sprmSD"] },
  { name: "Tables", patterns: ["plcftap", "tableRow"] },
  { name: "Headers/footers streams", patterns: ["plcfhddText", "headerStream"] },
  { name: "Pictures (Escher/Data)", patterns: ["fescher", "pictureData"] },
];

const XLS_FEATURES: Feature[] = [
  { name: "Workbook globals/BOF", patterns: ["0x0809", "BOF"] },
  { name: "Bound sheets (BOUNDSHEET)", patterns: ["0x0085", "boundSheet", "BOUNDSHEET"] },
  { name: "Shared strings (SST)", patterns: ["0x00FC", "sst"] },
  { name: "Number cells (NUMBER)", patterns: ["0x0203"] },
  { name: "RK/MulRK numeric cells", patterns: ["0x027E", "0x00BD", "rk"] },
  { name: "Label/LabelSst text cells", patterns: ["0x00FD", "0x0204", "label"] },
  { name: "RString cells", patterns: ["0x00D6"] },
  { name: "Bool/error cells", patterns: ["0x0205"] },
  { name: "Blank cells", patterns: ["0x0201"] },
  { name: "Formula cells + cached values", patterns: ["0x0006", "formula"] },
  { name: "Shared formulas (SHRFMLA)", patterns: ["0x04BC", "shrfmla"] },
  { name: "Array formulas (ARRAY)", patterns: ["0x0221"] },
  { name: "XF style records", patterns: ["0x00E0", "xf"] },
  { name: "Number formats (FORMAT)", patterns: ["0x041e"] },
  { name: "Fonts (FONT)", patterns: ["0x0031"] },
  { name: "Merged cells (MERGEDCELLS)", patterns: ["0x00e5"] },
  { name: "Column width/info (COLINFO)", patterns: ["0x007d"] },
  { name: "Row height (ROW)", patterns: ["0x0208"] },
  { name: "Defined names (NAME)", patterns: ["0x0018", "definedName"] },
  { name: "External references (SUPBOOK/EXTERNSHEET)", patterns: ["0x0017", "0x01AE", "supbook"] },
  { name: "Hyperlinks (HLINK)", patterns: ["0x01B8", "hlink"] },
  { name: "Notes/comments (NOTE/TXO)", patterns: ["0x001C", "0x01B6", "txo"] },
  { name: "Drawings (MSODRAWING)", patterns: ["0x00EC", "0x00EB"] },
  { name: "Conditional formatting (CONDFMT/CF)", patterns: ["0x01B0", "0x01B1", "condfmt"] },
  { name: "Page setup", patterns: ["0x00A1", "pagesetup"] },
  { name: "Codepage/encoding", patterns: ["0x0042"] },
  { name: "Date mode (1904)", patterns: ["0x0022"] },
  { name: "Palette (PALETTE)", patterns: ["0x0092"] },
];

const PPT_FEATURES: Feature[] = [
  { name: "Current user atom", patterns: ["4086", "currentUserAtom", "CurrentUserAtom"] },
  { name: "User edit atom", patterns: ["4085", "userEditAtom"] },
  { name: "Persist pointer blocks", patterns: ["6001", "6002", "persistPointer"] },
  { name: "Document container", patterns: ["1000", "document"] },
  { name: "Slides (SlideContainer)", patterns: ["1006"] },
  { name: "Slide atoms", patterns: ["1007"] },
  { name: "Main masters", patterns: ["1016", "mainMaster"] },
  {
    name: "Text extraction (TextChars/TextBytes)",
    patterns: ["4000", "4008", "textChars", "textBytes"],
  },
  { name: "Text header/formatting", patterns: ["3999", "textHeader"] },
  { name: "Slide list with text", patterns: ["4080"] },
  { name: "Escher shapes", patterns: ["0xf00a", "0xf004", "escherShape"] },
  { name: "Escher shape properties", patterns: ["0xf00b", "escherShapeProperties"] },
  { name: "Escher pictures (BSE/BStore)", patterns: ["0xf001", "0xf007", "escherBse"] },
  {
    name: "Escher groups/anchors",
    patterns: ["0xf009", "0xf00f", "0xf010", "escherChildAnchor", "escherClientAnchor"],
  },
  { name: "Textboxes (ClientTextbox)", patterns: ["0xf00d", "escherClientTextbox"] },
  { name: "Drawing containers (PPDrawing)", patterns: ["1036", "ppDrawing"] },
  { name: "Date/time atoms", patterns: ["4006"] },
  { name: "Slide number atoms", patterns: ["4056"] },
  { name: "Notes slides", patterns: ["1008", "notes"] },
  { name: "Animations", patterns: ["anim"] },
  { name: "Encryption (DocumentEncryptionAtom)", patterns: ["12052"] },
  { name: "Outline text references", patterns: ["3998", "outlineTextReference"] },
];

const FORMATS: Array<[string, string[], Feature[]]> = [
  ["doc", ["packages/doc/src"], DOC_FEATURES],
  ["xls", ["packages/xls/src"], XLS_FEATURES],
  ["ppt", ["packages/ppt/src"], PPT_FEATURES],
];

function main() {
  const args = process.argv.slice(2);
  const showMissing = !args.includes("--summary");
  const filter = args.find((arg) => !arg.startsWith("--"));

  console.log("======================================================================");
  console.log("Legacy Binary Format Coverage");
  console.log("======================================================================");
  console.log("| Format | Features     | Coverage |");
  console.log("|--------|--------------|----------|");

  const missingByFormat: Array<[string, Feature[]]> = [];
  let totalCovered = 0;
  let totalFeatures = 0;

  for (const [format, dirs, features] of FORMATS) {
    if (filter && format !== filter) continue;
    let code = "";
    for (const dir of dirs) {
      const absDir = path.resolve(ROOT_DIR, dir);
      for (const file of fs.readdirSync(absDir)) {
        if (file.endsWith(".ts") && !file.endsWith(".spec.ts")) {
          code += fs.readFileSync(path.join(absDir, file), "utf8");
        }
      }
    }
    const covered = features.filter((feature) =>
      feature.patterns.some((pattern) => code.toLowerCase().includes(pattern.toLowerCase())),
    );
    const missing = features.filter((feature) => !covered.includes(feature));
    totalCovered += covered.length;
    totalFeatures += features.length;
    const percent = ((covered.length / features.length) * 100).toFixed(1);
    console.log(
      `| ${format.padEnd(6)} | ${(covered.length + "/" + features.length).padEnd(12)} | ${percent.padStart(8)}% |`,
    );
    missingByFormat.push([format, missing]);
  }

  console.log("|--------|--------------|----------|");
  const totalPercent = ((totalCovered / totalFeatures) * 100).toFixed(1);
  console.log(`| **TOTAL** | **${totalCovered}/${totalFeatures}** | **${totalPercent}%** |`);

  if (showMissing) {
    for (const [format, missing] of missingByFormat) {
      if (missing.length === 0) continue;
      console.log(`\nMissing ${format}: (${missing.length})`);
      for (const feature of missing) console.log(`  - ${feature.name}`);
    }
  }
}

main();
