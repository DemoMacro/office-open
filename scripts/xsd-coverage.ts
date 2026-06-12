/**
 * XSD Coverage Analysis Tool
 *
 * Parses OOXML Transitional XSD schemas, extracts all element/attribute declarations,
 * scans the codebase for implementations, and generates a coverage report.
 *
 * Detection strategy: strips comments, then uses targeted regex patterns to find
 * element/attribute names in specific XML construction and parsing contexts.
 * No broad heuristics — only matches names in provable implementation patterns.
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

/** Recursively collect all .ts files in a directory (non-spec, non-bench) */
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
      !entry.name.endsWith(".bench.ts") &&
      !entry.name.endsWith(".d.ts")
    ) {
      results.push(fullPath);
    }
  }
  return results;
}

/** Cache for stripped file contents: absolute path → comment-free content */
const strippedCache = new Map<string, string>();

/**
 * Strip JS/TS comments from source code while preserving string literals.
 * Handles: // line comments, /* block comments *​/, template literals.
 */
function stripComments(content: string): string {
  const len = content.length;
  const out: string[] = [];
  let i = 0;

  while (i < len) {
    const ch = content[i];

    // String literal — skip through intact
    if (ch === '"' || ch === "'") {
      const quote = ch;
      out.push(ch);
      i++;
      while (i < len && content[i] !== quote) {
        if (content[i] === "\\") {
          out.push(content[i++]);
        }
        if (i < len) out.push(content[i++]);
      }
      if (i < len) out.push(content[i++]); // closing quote
      continue;
    }

    // Template literal — skip through, handle nested ${...}
    if (ch === "`") {
      out.push(ch);
      i++;
      let depth = 0;
      while (i < len) {
        if (content[i] === "\\") {
          out.push(content[i++]);
          if (i < len) out.push(content[i++]);
          continue;
        }
        if (content[i] === "`" && depth === 0) {
          out.push(content[i++]);
          break;
        }
        if (content[i] === "$" && content[i + 1] === "{") {
          depth++;
          out.push(content[i++], content[i++]);
          continue;
        }
        if (content[i] === "}" && depth > 0) {
          depth--;
          out.push(content[i++]);
          continue;
        }
        out.push(content[i++]);
      }
      continue;
    }

    // Line comment
    if (ch === "/" && content[i + 1] === "/") {
      // Replace with spaces to preserve positions
      while (i < len && content[i] !== "\n") {
        out.push(" ");
        i++;
      }
      continue;
    }

    // Block comment
    if (ch === "/" && content[i + 1] === "*") {
      out.push(" ", " ");
      i += 2;
      while (i < len && !(content[i] === "*" && content[i + 1] === "/")) {
        out.push(content[i] === "\n" ? "\n" : " ");
        i++;
      }
      if (i < len) {
        out.push(" ", " "); // */
        i += 2;
      }
      continue;
    }

    out.push(ch);
    i++;
  }

  return out.join("");
}

function readFileStripped(filePath: string): string {
  const cached = strippedCache.get(filePath);
  if (cached !== undefined) return cached;

  const raw = fs.readFileSync(filePath, "utf-8");
  const stripped = stripComments(raw);
  strippedCache.set(filePath, stripped);
  return stripped;
}

// ── XSD Parsing ──

/**
 * Deprecated/removed OOXML elements to exclude from coverage analysis.
 * These are features Microsoft has removed from Office or disabled by default.
 */
const DEPRECATED_ELEMENTS = new Set([
  // Smart Tags (removed Office 2010)
  "cellSmartTag",
  "cellSmartTagPr",
  "cellSmartTags",
  "smartTagPr",
  "smartTagType",
  "smartTagTypes",
  "smartTags",
  // DDE (disabled by default, security risk)
  "ddeLink",
  "ddeItems",
  "ddeItem",
  "values", // CT_DdeValues — DDE-specific
  "val", // CT_DdeValue — DDE-specific element (not attribute)
]);

/** Elements only deprecated in SML (val appears in dml-chart, pml too) */
const DEPRECATED_ELEMENTS_SML_ONLY = new Set(["val"]);

/**
 * Deprecated/removed OOXML attributes to exclude from coverage analysis.
 */
const DEPRECATED_ATTRIBUTES = new Set([
  // Smart Tags (removed Office 2010)
  "xmlBased",
  "namespaceUri",
  // DDE (disabled by default, security risk)
  "ddeService",
  "ddeTopic",
  "ole",
  "preferPic",
  // OLAP-only attributes
  "bc",
  "fc",
  "in",
  "st",
  "un",
  "ct",
  "fi",
]);

/**
 * Attribute names too generic to reliably track in source code.
 * These are excluded from the coverage denominator — they don't count
 * as "implemented" OR "missing", they're simply unmeasurable.
 */
const UNTRACKABLE_ATTRS = new Set([
  // Ultra-common in both XML and JS contexts
  "val",
  "type",
  "name",
  "id",
  "value",
  "text",
  "count",
  "index",
  "style",
  "title",
  "lang",
  "class",
  "width",
  "height",
  "src",
  "href",
  "path",
  "format",
  "uri",
  "url",
  "ref",
  "min",
  "max",
  "size",
  "length",
  "step",
  "scheme",
  "port",
  "host",
  "xmlns",
  // Single-character names — impossible to attribute
  "a",
  "b",
  "c",
  "d",
  "e",
  "f",
  "g",
  "h",
  "i",
  "k",
  "m",
  "n",
  "o",
  "p",
  "r",
  "s",
  "t",
  "u",
  "v",
  "w",
  "x",
  "y",
  "z",
]);

function parseXsd(xsdPath: string): { elements: Set<string>; attributes: Set<string> } {
  const content = fs.readFileSync(xsdPath, "utf-8");
  const elements = new Set<string>();
  const attributes = new Set<string>();

  const isSml = xsdPath.includes("sml.xsd");
  const deprecated = isSml
    ? new Set([...DEPRECATED_ELEMENTS, ...DEPRECATED_ELEMENTS_SML_ONLY])
    : DEPRECATED_ELEMENTS;

  const elementRe = /<xsd:element\s+name="([^"]+)"/g;
  let match: RegExpExecArray | null;
  while ((match = elementRe.exec(content)) !== null) {
    if (!deprecated.has(match[1])) {
      elements.add(match[1]);
    }
  }

  const attrRe = /<xsd:attribute\s+name="([^"]+)"/g;
  while ((match = attrRe.exec(content)) !== null) {
    attributes.add(match[1]);
  }

  return { elements, attributes };
}

// ── Code scanning ──

/**
 * Extract element names found in XML construction and parsing code.
 * Uses targeted patterns — no broad property/key heuristics.
 */
function extractUsedElements(config: XsdConfig): Set<string> {
  const found = new Set<string>();

  for (const dir of config.searchDirs) {
    const absDir = path.resolve(ROOT_DIR, dir);
    const files = collectTsFiles(absDir);

    for (const file of files) {
      const src = readFileStripped(file);
      let m: RegExpExecArray | null;

      // Pattern 1: XML tags in template literals — `<prefix:name` or `<name`
      // Catches stringify: `<w:p>`, `<a:solidFill`, `</c:chart>`
      // Also matches `<name${...}` (bare element followed by interpolation)
      const tagRe = /<\/?([a-z]+:)?([a-zA-Z][a-zA-Z0-9]*)[\s>\/"$]/g;
      while ((m = tagRe.exec(src)) !== null) {
        found.add(m[2]);
      }

      // Pattern 2: findChild(el, "prefix:name") — parse path
      const findChildRe = /findChild\([^,]+,\s*"([a-z]+:)?([a-zA-Z][a-zA-Z0-9]+)"/g;
      while ((m = findChildRe.exec(src)) !== null) {
        found.add(m[2]);
      }

      // Pattern 3: el.name === "prefix:name" or child.name === "prefix:name" — parse path
      const nameCmpRe = /\.name\s*[!=]==?\s*"([a-z]+:)?([a-zA-Z][a-zA-Z0-9]+)"/g;
      while ((m = nameCmpRe.exec(src)) !== null) {
        found.add(m[2]);
      }

      // Pattern 4: case "prefix:name": — switch in parse path
      const caseRe = /case\s+"([a-z]+:)?([a-zA-Z][a-zA-Z0-9]+)"/g;
      while ((m = caseRe.exec(src)) !== null) {
        found.add(m[2]);
      }

      // Pattern 5: element("prefix:name", ...) — descriptor builder / xml helper
      // Also catches selfCloseElement("name", ...), openElement("name", ...), closeElement("name")
      const elementCallRe =
        /(?:^|[.\s(])(?:selfClose|open|close)?Element\(\s*"([a-z]+:)?([a-zA-Z][a-zA-Z0-9]+)"/gm;
      while ((m = elementCallRe.exec(src)) !== null) {
        found.add(m[2]);
      }

      // Pattern 6: .child(k, "prefix:name", ...) / .children(k, "prefix:name", ...)
      const childCallRe = /\.child(?:ren)?\([^,]+,\s*"([a-z]+:)?([a-zA-Z][a-zA-Z0-9]+)"/g;
      while ((m = childCallRe.exec(src)) !== null) {
        found.add(m[2]);
      }

      // Pattern 7: valEl("prefix:name", ...) helper
      const valElRe = /valEl\(\s*"([a-z]+:)?([a-zA-Z][a-zA-Z0-9]+)"/g;
      while ((m = valElRe.exec(src)) !== null) {
        found.add(m[2]);
      }
    }
  }

  // ── Dynamic template tracking ──
  // When a file uses `<prefix:${var}>` to generate XML elements dynamically,
  // extract string literal values from the same file that could be the variable's value.
  // This catches patterns like `<p:${type}/>` where type is a union of string literals.

  // Collect files that have dynamic element patterns
  const dynamicFiles = new Set<string>();
  for (const dir of config.searchDirs) {
    const absDir = path.resolve(ROOT_DIR, dir);
    const files = collectTsFiles(absDir);
    for (const file of files) {
      const src = readFileStripped(file);
      if (/<[a-z]+:\$\{[a-zA-Z]/.test(src) || /<\$\{[a-zA-Z]/.test(src)) {
        dynamicFiles.add(file);
      }
    }
  }

  // For each dynamic file, extract string literal values that could be element names
  if (dynamicFiles.size > 0) {
    const strLitRe = /"([a-zA-Z][a-zA-Z0-9]+)"/g;
    for (const file of dynamicFiles) {
      const src = readFileStripped(file);
      let m: RegExpExecArray | null;
      while ((m = strLitRe.exec(src)) !== null) {
        // Only add names that look like XML element names (not JS identifiers like "string", "number")
        found.add(m[1]);
      }
    }
  }

  return found;
}

/**
 * Extract attribute names found in XML construction and parsing code.
 * Only matches attributes in provable XML/parsing contexts.
 */
function extractUsedAttributes(config: XsdConfig): Set<string> {
  const found = new Set<string>();

  for (const dir of config.searchDirs) {
    const absDir = path.resolve(ROOT_DIR, dir);
    const files = collectTsFiles(absDir);

    for (const file of files) {
      const src = readFileStripped(file);
      let m: RegExpExecArray | null;

      // Pattern 1: XML attribute in template — name=, name="${, name='
      // e.g. w:val="${v}", sz="1200", xmlns:w="..."
      const xmlAttrRe = /\b([a-zA-Z][a-zA-Z0-9]+)=["'`]/g;
      while ((m = xmlAttrRe.exec(src)) !== null) {
        found.add(m[1]);
      }

      // Pattern 2: el.attributes["name"] or el.attributes["prefix:name"] — parse path
      const attrAccessRe = /\.attributes\[\s*["']([a-zA-Z][a-zA-Z0-9]+)["']\s*\]/g;
      while ((m = attrAccessRe.exec(src)) !== null) {
        found.add(m[1]);
      }

      // Pattern 3: bracket access with prefix — attrs["w:val"], result["w:spacing"], etc.
      // This catches stringify (attrs["w:name"]) and parse (result["w:name"]) patterns
      const bracketPrefixRe = /\[\s*["'][a-z]+:([a-zA-Z][a-zA-Z0-9]+)["']\s*\]/g;
      while ((m = bracketPrefixRe.exec(src)) !== null) {
        found.add(m[1]);
      }

      // Pattern 3: .attr(key, "xmlName") — descriptor builder
      const builderAttrRe = /\.attr\([^,]+,\s*["']([a-zA-Z][a-zA-Z0-9]+)["']/g;
      while ((m = builderAttrRe.exec(src)) !== null) {
        found.add(m[1]);
      }

      // Pattern 4: attr(el, "name") — parse helper from @office-open/xml
      const attrHelperRe = /\battr\([^,]+,\s*["']([a-zA-Z][a-zA-Z0-9]+)["']/g;
      while ((m = attrHelperRe.exec(src)) !== null) {
        found.add(m[1]);
      }

      // Pattern 5: attr(el, "prefix:name") — prefixed version
      const prefixedHelperRe = /\battr\([^,]+,\s*["'][a-z]+:([a-zA-Z][a-zA-Z0-9]+)["']/g;
      while ((m = prefixedHelperRe.exec(src)) !== null) {
        found.add(m[1]);
      }

      // Pattern 6: Dynamic template patterns — `}:name` (template interpolation suffix)
      const tmplSuffixRe = /\}:([a-zA-Z][a-zA-Z0-9]+)/g;
      while ((m = tmplSuffixRe.exec(src)) !== null) {
        found.add(m[1]);
      }
    }
  }

  return found;
}

/**
 * Cross-namespace fallback: some elements defined in one XSD namespace appear
 * in code under a different prefix (e.g. a:cNvPicPr → p:cNvPicPr).
 */
function findCrossNamespace(
  bareNames: readonly string[],
  searchDirs: readonly string[],
): Set<string> {
  if (bareNames.length === 0) return new Set();
  const found = new Set<string>();
  const nameSet = new Set(bareNames);
  const re = /[a-z]+:([a-zA-Z][a-zA-Z0-9]+)/g;

  for (const dir of searchDirs) {
    const absDir = path.resolve(ROOT_DIR, dir);
    const files = collectTsFiles(absDir);
    for (const file of files) {
      const src = readFileStripped(file);
      let m: RegExpExecArray | null;
      while ((m = re.exec(src)) !== null) {
        if (nameSet.has(m[1])) {
          found.add(m[1]);
        }
      }
    }
  }
  return found;
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
  untrackableAttributes: number;
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
      untrackableAttributes: 0,
    };
  }

  const { elements, attributes } = parseXsd(xsdPath);

  // ── Element coverage ──
  const usedElements = extractUsedElements(config);
  const elementNames = [...elements].sort();
  let missingElements = elementNames.filter((e) => !usedElements.has(e));

  // Cross-namespace fallback
  if (missingElements.length > 0 && config.searchMode === "prefix") {
    const crossNs = findCrossNamespace(missingElements, config.searchDirs);
    if (crossNs.size > 0) {
      missingElements = missingElements.filter((e) => !crossNs.has(e));
    }
  }

  // ── Attribute coverage ──
  const usedAttributes = extractUsedAttributes(config);
  const attrNames = [...attributes].sort();

  let untrackableCount = 0;
  let deprecatedCount = 0;
  const trackableAttrs = attrNames.filter((a) => {
    if (DEPRECATED_ATTRIBUTES.has(a)) {
      deprecatedCount++;
      return false;
    }
    if (UNTRACKABLE_ATTRS.has(a)) {
      untrackableCount++;
      return false;
    }
    return true;
  });

  const missingAttributes = trackableAttrs.filter((a) => !usedAttributes.has(a));
  const totalTrackable = trackableAttrs.length;

  return {
    config,
    totalElements: elementNames.length,
    implementedElements: elementNames.length - missingElements.length,
    missingElements,
    totalAttributes: totalTrackable,
    implementedAttributes: totalTrackable - missingAttributes.length,
    missingAttributes,
    untrackableAttributes: untrackableCount,
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
    untrackableAttributes,
  } = result;

  const elPct =
    totalElements > 0 ? ((implementedElements / totalElements) * 100).toFixed(1) : "0.0";
  const attrPct =
    totalAttributes > 0 ? ((implementedAttributes / totalAttributes) * 100).toFixed(1) : "0.0";

  const lines: string[] = [];
  lines.push(`=== ${config.xsdFile} (${config.description}) ===`);
  lines.push(`Elements: ${implementedElements}/${totalElements} implemented (${elPct}%)`);

  if (showMissing && missingElements.length > 0) {
    lines.push(`Missing elements (${missingElements.length}):`);
    for (const el of missingElements) {
      const prefix = config.searchMode === "bare" ? "" : config.prefix;
      lines.push(`  - ${prefix}${el}`);
    }
  }

  lines.push(
    `Attributes: ${implementedAttributes}/${totalAttributes} implemented (${attrPct}%)` +
      (untrackableAttributes > 0 ? ` (${untrackableAttributes} untrackable)` : ""),
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
  const header = "| XSD | Elements | El% | Attributes | Attr% | Untrackable |";
  const separator = "|-----|----------|-----|------------|-------|-------------|";
  const rows = results.map((r) => {
    const elPct =
      r.totalElements > 0 ? ((r.implementedElements / r.totalElements) * 100).toFixed(1) : "0.0";
    const attrPct =
      r.totalAttributes > 0
        ? ((r.implementedAttributes / r.totalAttributes) * 100).toFixed(1)
        : "0.0";
    const untr = r.untrackableAttributes > 0 ? String(r.untrackableAttributes) : "-";
    return `| ${r.config.label} | ${r.implementedElements}/${r.totalElements} | ${elPct}% | ${r.implementedAttributes}/${r.totalAttributes} | ${attrPct}% | ${untr} |`;
  });

  const totalEl = results.reduce((s, r) => s + r.totalElements, 0);
  const implEl = results.reduce((s, r) => s + r.implementedElements, 0);
  const totalAttr = results.reduce((s, r) => s + r.totalAttributes, 0);
  const implAttr = results.reduce((s, r) => s + r.implementedAttributes, 0);
  const totalUntr = results.reduce((s, r) => s + r.untrackableAttributes, 0);
  const totalElPct = totalEl > 0 ? ((implEl / totalEl) * 100).toFixed(1) : "0.0";
  const totalAttrPct = totalAttr > 0 ? ((implAttr / totalAttr) * 100).toFixed(1) : "0.0";

  rows.push(separator);
  rows.push(
    `| **TOTAL** | **${implEl}/${totalEl}** | **${totalElPct}%** | **${implAttr}/${totalAttr}** | **${totalAttrPct}%** | **${totalUntr}** |`,
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
        untrackable: r.untrackableAttributes,
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

  console.log("=".repeat(70));
  console.log("SUMMARY");
  console.log("=".repeat(70));
  console.log(formatSummaryTable(results));
}

main();
