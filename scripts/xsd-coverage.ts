/**
 * XSD Coverage Analysis Tool
 *
 * Parses OOXML Transitional XSD schemas, extracts all element/attribute declarations,
 * scans the codebase for implementations, and generates a coverage report.
 *
 * Strategy: reads source files directly with Node.js fs and uses regex to find
 * all used element/attribute names. Pure JS — no external process spawning.
 *
 * Usage:
 *   pnpm tsx scripts/xsd-coverage.ts              # full report
 *   pnpm tsx scripts/xsd-coverage.ts wml           # docx only
 *   pnpm tsx scripts/xsd-coverage.ts pml           # pptx only
 *   pnpm tsx scripts/xsd-coverage.ts sml           # xlsx only
 *   pnpm tsx scripts/xsd-coverage.ts dml-main      # DrawingML main
 *   pnpm tsx scripts/xsd-coverage.ts --missing     # show missing items (default)
 *   pnpm tsx scripts/xsd-coverage.ts --summary     # only show summary stats
 *   pnpm tsx scripts/xsd-coverage.ts --json        # JSON output
 */
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT_DIR = path.resolve(__dirname, "..");
const XSD_DIR = path.resolve(__dirname, "../ooxml-schemas/transitional");

// ── XSD → Code mapping configuration ──

interface XsdConfig {
  /** XSD file name (in ooxml-schemas/transitional/) */
  xsdFile: string;
  /** Short label for CLI filter and output */
  label: string;
  /** Human-readable description */
  description: string;
  /** Namespace prefix used in code (e.g. "w:", "p:", "a:") — empty for bare mode */
  prefix: string;
  /** Directories to search for implementations (relative to ROOT_DIR) */
  searchDirs: string[];
  /**
   * Search mode:
   * - "prefix" → elements appear as `"prefix:name"` in code (docx, pptx, core)
   * - "bare" → elements appear as `"name"` or `<name` in code (xlsx/sml)
   */
  searchMode: "prefix" | "bare";
}

const XSD_CONFIGS: XsdConfig[] = [
  {
    xsdFile: "wml.xsd",
    label: "wml",
    description: "WordprocessingML (docx)",
    prefix: "w:",
    searchDirs: ["packages/docx/src"],
    searchMode: "prefix",
  },
  {
    xsdFile: "pml.xsd",
    label: "pml",
    description: "PresentationML (pptx)",
    prefix: "p:",
    searchDirs: ["packages/pptx/src"],
    searchMode: "prefix",
  },
  {
    xsdFile: "sml.xsd",
    label: "sml",
    description: "SpreadsheetML (xlsx)",
    prefix: "",
    searchDirs: ["packages/xlsx/src"],
    searchMode: "bare",
  },
  {
    xsdFile: "dml-main.xsd",
    label: "dml-main",
    description: "DrawingML Main (core)",
    prefix: "a:",
    searchDirs: ["packages/core/src", "packages/docx/src", "packages/pptx/src"],
    searchMode: "prefix",
  },
  {
    xsdFile: "dml-chart.xsd",
    label: "dml-chart",
    description: "DrawingML Chart (core)",
    prefix: "c:",
    searchDirs: ["packages/core/src"],
    searchMode: "prefix",
  },
  {
    xsdFile: "dml-diagram.xsd",
    label: "dml-diagram",
    description: "DrawingML Diagram/SmartArt (core)",
    prefix: "dgm:",
    searchDirs: ["packages/core/src"],
    searchMode: "prefix",
  },
  {
    xsdFile: "dml-wordprocessingDrawing.xsd",
    label: "dml-wp",
    description: "DrawingML WordprocessingDrawing (docx)",
    prefix: "wp:",
    searchDirs: ["packages/docx/src", "packages/core/src"],
    searchMode: "prefix",
  },
  {
    xsdFile: "dml-spreadsheetDrawing.xsd",
    label: "dml-xdr",
    description: "DrawingML SpreadsheetDrawing (xlsx)",
    prefix: "xdr:",
    searchDirs: ["packages/xlsx/src", "packages/core/src"],
    searchMode: "prefix",
  },
  {
    xsdFile: "dml-picture.xsd",
    label: "dml-pic",
    description: "DrawingML Picture (core)",
    prefix: "pic:",
    searchDirs: ["packages/core/src", "packages/docx/src", "packages/pptx/src"],
    searchMode: "prefix",
  },
  {
    xsdFile: "dml-lockedCanvas.xsd",
    label: "dml-lc",
    description: "DrawingML LockedCanvas",
    prefix: "lc:",
    searchDirs: ["packages/core/src", "packages/docx/src", "packages/pptx/src"],
    searchMode: "prefix",
  },
  {
    xsdFile: "dml-chartDrawing.xsd",
    label: "dml-cdr",
    description: "DrawingML ChartDrawing",
    prefix: "cdr:",
    searchDirs: ["packages/core/src"],
    searchMode: "prefix",
  },
  {
    xsdFile: "vml-main.xsd",
    label: "vml",
    description: "VML Main",
    prefix: "v:",
    searchDirs: ["packages/docx/src", "packages/xlsx/src"],
    searchMode: "prefix",
  },
  {
    xsdFile: "vml-officeDrawing.xsd",
    label: "vml-office",
    description: "VML Office Drawing",
    prefix: "o:",
    searchDirs: ["packages/docx/src"],
    searchMode: "prefix",
  },
  {
    xsdFile: "vml-wordprocessingDrawing.xsd",
    label: "vml-word",
    description: "VML WordprocessingDrawing",
    prefix: "w10:",
    searchDirs: ["packages/docx/src"],
    searchMode: "prefix",
  },
  {
    xsdFile: "vml-spreadsheetDrawing.xsd",
    label: "vml-excel",
    description: "VML SpreadsheetDrawing",
    prefix: "x:",
    searchDirs: ["packages/xlsx/src"],
    searchMode: "prefix",
  },
  {
    xsdFile: "shared-math.xsd",
    label: "math",
    description: "Math (docx)",
    prefix: "m:",
    searchDirs: ["packages/docx/src"],
    searchMode: "prefix",
  },
];

// ── File reading utilities ──

/** Recursively collect all .ts files in a directory (non-spec files) */
function collectTsFiles(dir: string): string[] {
  const results: string[] = [];
  if (!fs.existsSync(dir)) return results;

  const entries = fs.readdirSync(dir, { withFileTypes: true });
  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      results.push(...collectTsFiles(fullPath));
    } else if (
      entry.isFile() &&
      entry.name.endsWith(".ts") &&
      !entry.name.endsWith(".spec.ts") &&
      !entry.name.endsWith(".bench.ts")
    ) {
      results.push(fullPath);
    }
  }
  return results;
}

/** Cache for file contents: absolute path → content */
const fileCache = new Map<string, string>();

function readFileCached(filePath: string): string {
  const cached = fileCache.get(filePath);
  if (cached !== undefined) return cached;

  const content = fs.readFileSync(filePath, "utf-8");
  fileCache.set(filePath, content);
  return content;
}

// ── XSD Parsing ──

function parseXsd(xsdPath: string): { elements: Set<string>; attributes: Set<string> } {
  const content = fs.readFileSync(xsdPath, "utf-8");
  const elements = new Set<string>();
  const attributes = new Set<string>();

  const elementRe = /<xsd:element\s+name="([^"]+)"/g;
  let match: RegExpExecArray | null;
  while ((match = elementRe.exec(content)) !== null) {
    elements.add(match[1]);
  }

  const attrRe = /<xsd:attribute\s+name="([^"]+)"/g;
  while ((match = attrRe.exec(content)) !== null) {
    attributes.add(match[1]);
  }

  return { elements, attributes };
}

// ── Code scanning (pure JS, no shell) ──

/**
 * Scan source directories and extract all element/attribute names used in code.
 * For prefix mode: extracts names matching "prefix:word" patterns in source files.
 * Returns a Set of bare element names (without prefix).
 */
function extractUsedNamesFromCode(config: XsdConfig): Set<string> {
  const found = new Set<string>();
  const prefixStr = config.prefix;

  for (const dir of config.searchDirs) {
    const absDir = path.resolve(ROOT_DIR, dir);
    const files = collectTsFiles(absDir);

    for (const file of files) {
      const content = readFileCached(file);

      if (config.searchMode === "prefix" && prefixStr) {
        // Pattern 1: literal "prefix:word" strings (e.g., "w:p", "a:solidFill")
        const re = new RegExp(`${escapeRegExp(prefixStr)}[a-zA-Z][a-zA-Z0-9]*`, "g");
        let m: RegExpExecArray | null;
        while ((m = re.exec(content)) !== null) {
          const bareName = m[0].slice(prefixStr.length);
          if (bareName) found.add(bareName);
        }

        // Pattern 2: template string dynamic prefix like `${element}:spPr`
        // Matches `}:word` which indicates dynamic prefix + static suffix
        const tmplRe = /\}(:[a-zA-Z][a-zA-Z0-9]*)/g;
        while ((m = tmplRe.exec(content)) !== null) {
          // m[1] is like ":spPr" — the suffix
          const suffix = m[1].slice(1);
          if (suffix) found.add(suffix);
        }
      } else {
        // bare mode (xlsx/sml): match "word" in string literals and <word in XML strings
        // SML has many single-char elements (b, c, d, e, f, i, k, m, n, p, r, s, t, u, v, x)
        // so the regex allows single-char names in XML tag context only (not quoted strings,
        // where single-char matches would produce too many false positives).
        const quotedRe = /"([a-zA-Z][a-zA-Z0-9]+)"/g;
        let m: RegExpExecArray | null;
        while ((m = quotedRe.exec(content)) !== null) {
          found.add(m[1]);
        }
        const tagRe = /<([a-zA-Z][a-zA-Z0-9]*)[\s>\/"]/g;
        while ((m = tagRe.exec(content)) !== null) {
          found.add(m[1]);
        }
      }
    }
  }

  return found;
}

function escapeRegExp(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// ── Report generation ──

interface CoverageResult {
  config: XsdConfig;
  totalElements: number;
  implementedElements: number;
  missingElements: string[];
  totalAttributes: number;
  implementedAttributes: number;
  missingAttributes: string[];
}

function analyzeXsd(config: XsdConfig): CoverageResult {
  const xsdPath = path.resolve(XSD_DIR, config.xsdFile);
  if (!fs.existsSync(xsdPath)) {
    console.error(`XSD not found: ${xsdPath}`);
    return {
      config,
      totalElements: 0,
      implementedElements: 0,
      missingElements: [],
      totalAttributes: 0,
      implementedAttributes: 0,
      missingAttributes: [],
    };
  }

  const { elements, attributes } = parseXsd(xsdPath);
  const usedNames = extractUsedNamesFromCode(config);

  // Elements: a name is "implemented" if it appears anywhere in source
  const elementNames = [...elements].sort();
  const missingElements = elementNames.filter((e) => !usedNames.has(e));

  // Attributes: search for bare attribute names using diverse code patterns.
  // Unlike elements (which consistently appear as "prefix:name" strings),
  // attributes appear as bare object keys, property accesses, template attrs, etc.
  const attrNames = [...attributes].sort();
  const usedAttrNames = new Set(usedNames);

  for (const dir of config.searchDirs) {
    const absDir = path.resolve(ROOT_DIR, dir);
    const files = collectTsFiles(absDir);
    for (const file of files) {
      const content = readFileCached(file);
      let m: RegExpExecArray | null;

      // Object key: allowOverlap: "1", showGridLines: false
      // Also matches TypeScript optional property syntax: noSelect?:
      const objKeyRe = /\b([a-zA-Z][a-zA-Z0-9]{1,})\??\s*:/g;
      while ((m = objKeyRe.exec(content)) !== null) {
        usedAttrNames.add(m[1]);
      }

      // Property access: svMap.showGridLines, obj.applyFont
      const propRe = /\.([a-zA-Z][a-zA-Z0-9]{1,})/g;
      while ((m = propRe.exec(content)) !== null) {
        usedAttrNames.add(m[1]);
      }

      // XML attribute in template literal: showGridLines="${...}"
      const xmlAttrRe = /\b([a-zA-Z][a-zA-Z0-9]+)=["`]/g;
      while ((m = xmlAttrRe.exec(content)) !== null) {
        usedAttrNames.add(m[1]);
      }

      // BuilderElement key mapping: key: "bld", key: "prevAc", etc.
      // This catches XSD attribute names used as XML key values in BuilderElement attributes.
      const keyValueRe = /key:\s*["']([a-zA-Z][a-zA-Z0-9]+)["']/g;
      while ((m = keyValueRe.exec(content)) !== null) {
        usedAttrNames.add(m[1]);
      }

      // xmlKey mapping in lookup arrays: xmlKey: "verticies", xmlKey: "adjusthandles", etc.
      // Used by OLock boolMap and similar patterns where key values are indirect.
      const xmlKeyRe = /xmlKey:\s*["']([a-zA-Z][a-zA-Z0-9]+)["']/g;
      while ((m = xmlKeyRe.exec(content)) !== null) {
        usedAttrNames.add(m[1]);
      }

      // For prefix mode: "prefix:name" in attribute objects (e.g., { _attr: { "w:val": ... } })
      if (config.searchMode === "prefix" && config.prefix) {
        const prefixAttrRe = new RegExp(
          `["']${escapeRegExp(config.prefix)}([a-zA-Z][a-zA-Z0-9]+)["']`,
          "g",
        );
        while ((m = prefixAttrRe.exec(content)) !== null) {
          usedAttrNames.add(m[1]);
        }
      }
    }
  }

  // Generic attribute names that are almost certainly handled somewhere
  const genericAttrs = new Set([
    "val",
    "type",
    "name",
    "id",
    "value",
    "count",
    "index",
    "text",
    "l",
    "b",
    "style",
    "title",
    "lang",
    "class",
    "width",
    "height",
    "src",
    "href",
    "xmlns",
    "scheme",
    "port",
    "host",
    "path",
    "format",
    "uri",
    "url",
    "r",
    "c",
    "t",
    "s",
    "d",
    "g",
    "h",
    "i",
    "n",
    "u",
    "w",
    "x",
    "y",
    "z",
    "ref",
    "min",
    "max",
    "step",
    "size",
    "length",
    "numFmtId",
    "fontId",
    "fillId",
    "borderId",
    "xfId",
    "v",
    "UpdateMode",
    "a",
    "o",
  ]);

  const missingAttributes = attrNames.filter((a) => {
    if (usedAttrNames.has(a)) return false;
    if (genericAttrs.has(a)) return false;
    return true;
  });

  return {
    config,
    totalElements: elementNames.length,
    implementedElements: elementNames.length - missingElements.length,
    missingElements,
    totalAttributes: attrNames.length,
    implementedAttributes: attrNames.length - missingAttributes.length,
    missingAttributes,
  };
}

function formatReport(result: CoverageResult, showMissing: boolean): string {
  const {
    config,
    totalElements,
    implementedElements,
    missingElements,
    totalAttributes,
    implementedAttributes,
    missingAttributes,
  } = result;

  const elPct =
    totalElements > 0 ? ((implementedElements / totalElements) * 100).toFixed(1) : "0.0";
  const attrPct =
    totalAttributes > 0 ? ((implementedAttributes / totalAttributes) * 100).toFixed(1) : "0.0";

  const lines: string[] = [];
  lines.push(`=== ${config.xsdFile} (${config.description}) ===`);
  lines.push(`Elements: ${totalElements} total, ${implementedElements} implemented (${elPct}%)`);

  if (showMissing && missingElements.length > 0) {
    lines.push(`Missing elements (${missingElements.length}):`);
    for (const el of missingElements) {
      const prefix = config.searchMode === "bare" ? "" : config.prefix;
      lines.push(`  - ${prefix}${el}`);
    }
  }

  lines.push(
    `Attributes: ${totalAttributes} total, ${implementedAttributes} implemented (${attrPct}%)`,
  );

  if (showMissing && missingAttributes.length > 0) {
    const shown = missingAttributes.slice(0, 50);
    lines.push(
      `Missing attributes (${missingAttributes.length}${missingAttributes.length > 50 ? ", showing first 50" : ""}):`,
    );
    for (const attr of shown) {
      lines.push(`  - ${attr}`);
    }
  }

  lines.push("");
  return lines.join("\n");
}

function formatSummaryTable(results: CoverageResult[]): string {
  const header = "| XSD | Elements | % | Attributes | % |";
  const separator = "|-----|----------|---|------------|---|";
  const rows = results.map((r) => {
    const elPct =
      r.totalElements > 0 ? ((r.implementedElements / r.totalElements) * 100).toFixed(1) : "0.0";
    const attrPct =
      r.totalAttributes > 0
        ? ((r.implementedAttributes / r.totalAttributes) * 100).toFixed(1)
        : "0.0";
    return `| ${r.config.label} | ${r.implementedElements}/${r.totalElements} | ${elPct}% | ${r.implementedAttributes}/${r.totalAttributes} | ${attrPct}% |`;
  });

  const totalEl = results.reduce((s, r) => s + r.totalElements, 0);
  const implEl = results.reduce((s, r) => s + r.implementedElements, 0);
  const totalAttr = results.reduce((s, r) => s + r.totalAttributes, 0);
  const implAttr = results.reduce((s, r) => s + r.implementedAttributes, 0);
  const totalElPct = totalEl > 0 ? ((implEl / totalEl) * 100).toFixed(1) : "0.0";
  const totalAttrPct = totalAttr > 0 ? ((implAttr / totalAttr) * 100).toFixed(1) : "0.0";

  rows.push(separator);
  rows.push(
    `| **TOTAL** | **${implEl}/${totalEl}** | **${totalElPct}%** | **${implAttr}/${totalAttr}** | **${totalAttrPct}%** |`,
  );

  return [header, separator, ...rows].join("\n");
}

// ── Main ──

function main() {
  const args = process.argv.slice(2);
  const showMissing = !args.includes("--summary");
  const jsonOutput = args.includes("--json");
  const filters = args.filter((a) => !a.startsWith("--"));

  let configs = XSD_CONFIGS;
  if (filters.length > 0) {
    configs = XSD_CONFIGS.filter((c) => filters.includes(c.label) || filters.includes(c.xsdFile));
    if (configs.length === 0) {
      console.error(
        `No matching XSD configs. Available labels: ${XSD_CONFIGS.map((c) => c.label).join(", ")}`,
      );
      process.exit(1);
    }
  }

  console.log(`Analyzing ${configs.length} XSD schema(s)...\n`);

  const results: CoverageResult[] = [];
  for (const config of configs) {
    process.stderr.write(`  Scanning ${config.xsdFile}...`);
    const result = analyzeXsd(config);
    results.push(result);
    const elPct =
      result.totalElements > 0
        ? ((result.implementedElements / result.totalElements) * 100).toFixed(1)
        : "0.0";
    process.stderr.write(
      ` ${result.implementedElements}/${result.totalElements} elements (${elPct}%)\n`,
    );
  }

  if (jsonOutput) {
    const jsonResults = results.map((r) => ({
      xsd: r.config.xsdFile,
      label: r.config.label,
      description: r.config.description,
      elements: {
        total: r.totalElements,
        implemented: r.implementedElements,
        coverage:
          r.totalElements > 0 ? +((r.implementedElements / r.totalElements) * 100).toFixed(1) : 0,
        missing: r.missingElements,
      },
      attributes: {
        total: r.totalAttributes,
        implemented: r.implementedAttributes,
        coverage:
          r.totalAttributes > 0
            ? +((r.implementedAttributes / r.totalAttributes) * 100).toFixed(1)
            : 0,
        missing: r.missingAttributes,
      },
    }));
    console.log(JSON.stringify(jsonResults, null, 2));
    return;
  }

  for (const result of results) {
    console.log(formatReport(result, showMissing));
  }

  console.log("=".repeat(60));
  console.log("SUMMARY");
  console.log("=".repeat(60));
  console.log(formatSummaryTable(results));
}

main();
