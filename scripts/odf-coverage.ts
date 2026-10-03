/**
 * ODF Relax NG Coverage Analysis Tool
 *
 * Parses the ODF 1.3 Relax NG schemas, extracts all element declarations,
 * scans the ODF package for implementations, and generates a coverage report.
 *
 * Detection strategy mirrors xsd-coverage.ts: elements are implemented when
 * their prefixed name appears in XML construction/parsing contexts
 * (`"text:p"` string literals or `<text:p` template literals).
 *
 * Usage:
 *   pnpm tsx scripts/odf-coverage.ts             # full report
 *   pnpm tsx scripts/odf-coverage.ts text        # one prefix only
 *   pnpm tsx scripts/odf-coverage.ts --missing   # show missing elements (default)
 *   pnpm tsx scripts/odf-coverage.ts --summary   # only summary table
 */
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT_DIR = path.resolve(__dirname, "..");
const SEARCH_DIRS = ["packages/odf/src"];

const SCHEMA_FILES: Array<[string, string]> = [
  ["odf-schemas/OpenDocument-v1.3-schema.rng", "odf"],
  ["odf-schemas/OpenDocument-v1.3-manifest-schema.rng", "manifest"],
];

interface PrefixStats {
  elements: string[];
  covered: string[];
}

function loadCode(): string {
  let code = "";
  for (const dir of SEARCH_DIRS) {
    const absDir = path.resolve(ROOT_DIR, dir);
    for (const file of fs.readdirSync(absDir)) {
      if (file.endsWith(".ts") && file !== "odf.spec.ts") {
        code += fs.readFileSync(path.join(absDir, file), "utf8");
      }
    }
  }
  return code;
}

function isImplemented(code: string, element: string): boolean {
  return code.includes(`"${element}"`) || code.includes(`<${element}`);
}

function main() {
  const args = process.argv.slice(2);
  const showMissing = !args.includes("--summary");
  const prefixFilter = args.find((arg) => !arg.startsWith("--"));

  const code = loadCode();
  const prefixes = new Map<string, PrefixStats>();

  for (const [schemaFile] of SCHEMA_FILES) {
    const rng = fs.readFileSync(path.resolve(ROOT_DIR, schemaFile), "utf8");
    for (const match of rng.matchAll(/<rng:element name="([^"]+)"/g)) {
      const element = match[1]!;
      const prefix = element.split(":")[0] ?? "";
      const stats = prefixes.get(prefix) ?? { elements: [], covered: [] };
      if (!stats.elements.includes(element)) stats.elements.push(element);
      if (isImplemented(code, element) && !stats.covered.includes(element))
        stats.covered.push(element);
      prefixes.set(prefix, stats);
    }
  }

  const rows = [...prefixes.entries()]
    .filter(([prefix]) => !prefixFilter || prefix === prefixFilter)
    .sort((a, b) => b[1].elements.length - a[1].elements.length);

  let totalElements = 0;
  let totalCovered = 0;
  for (const [, stats] of rows) {
    totalElements += stats.elements.length;
    totalCovered += stats.covered.length;
  }

  console.log("======================================================================");
  console.log("ODF Relax NG Coverage");
  console.log("======================================================================");
  console.log("| Prefix      | Elements      | Coverage |");
  console.log("|-------------|---------------|----------|");
  for (const [prefix, stats] of rows) {
    const percent = ((stats.covered.length / stats.elements.length) * 100).toFixed(1);
    console.log(
      `| ${prefix.padEnd(11)} | ${(stats.covered.length + "/" + stats.elements.length).padEnd(13)} | ${percent.padStart(8)}% |`,
    );
  }
  console.log("|-------------|---------------|----------|");
  const totalPercent = ((totalCovered / totalElements) * 100).toFixed(1);
  console.log(`| **TOTAL**   | **${totalCovered}/${totalElements}** | **${totalPercent}%** |`);

  if (showMissing) {
    for (const [prefix, stats] of rows) {
      const missing = stats.elements.filter((element) => !stats.covered.includes(element));
      if (missing.length === 0) continue;
      console.log(`\nMissing ${prefix}: (${missing.length})`);
      for (const element of missing) console.log(`  - ${element}`);
    }
  }
}

main();
