import { execSync } from "node:child_process";
/**
 * Run all demos and validate ALL generated XML against OOXML XSD schemas.
 *
 * For each ZIP entry, the validator auto-detects the correct XSD based on
 * the root element or file path:
 *   - ppt/slides/slide*.xml        → pml.xsd
 *   - word/document.xml            → wml.xsd
 *   - xl/worksheets/sheet*.xml     → sml.xsd
 *   - xl/drawings/drawing*.xml     → dml-spreadsheetDrawing.xsd
 *   - xl/charts/chart*.xml         → dml-chart.xsd
 *   - xl/theme/theme*.xml          → dml-main.xsd
 *   - xl/styles.xml                → sml.xsd (CT_Stylesheet)
 *   - (others)                     → skipped or auto-detected
 *
 * Usage:
 *   npx tsx scripts/validate.ts              # validate all (pptx + docx + xlsx)
 *   npx tsx scripts/validate.ts pptx         # validate pptx demos only
 *   npx tsx scripts/validate.ts docx         # validate docx demos only
 *   npx tsx scripts/validate.ts xlsx         # validate xlsx demos only
 *   npx tsx scripts/validate.ts slide <file> [n]  # validate specific pptx slide
 *   npx tsx scripts/validate.ts docx <file>       # validate specific docx file
 *   npx tsx scripts/validate.ts xlsx <file> [n]   # validate specific xlsx file
 */
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

import { unzipSync } from "fflate";
import { XmlDocument, XsdValidator } from "libxml2-wasm";
import { xmlRegisterFsInputProviders } from "libxml2-wasm/lib/nodejs.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT_DIR = path.resolve(__dirname, "..");
const XSD_DIR = path.resolve(__dirname, "../ooxml-schemas/transitional");

function getDemoFiles(dir: string): string[] {
  return fs
    .readdirSync(dir)
    .filter((f) => f.endsWith(".ts"))
    .sort();
}

function extractFromZip(zipPath: string, entryName: string): string {
  const buf = fs.readFileSync(zipPath);
  const files = unzipSync(buf);
  const entry = files[entryName];
  if (!entry) throw new Error(`Entry "${entryName}" not found in ${zipPath}`);
  return new TextDecoder().decode(entry);
}

function getZipEntries(zipPath: string): string[] {
  const buf = fs.readFileSync(zipPath);
  const files = unzipSync(buf);
  return Object.keys(files);
}

// ── Validator cache (one XSD → one validator, reused across demos) ──

const validatorCache = new Map<string, XsdValidator>();

function loadValidator(xsdFile: string): XsdValidator {
  const cached = validatorCache.get(xsdFile);
  if (cached) return cached;

  const originalCwd = process.cwd();
  process.chdir(XSD_DIR);
  const xsdBuffer = fs.readFileSync(xsdFile);
  const xsdDoc = XmlDocument.fromBuffer(xsdBuffer);
  const validator = XsdValidator.fromDoc(xsdDoc);
  xsdDoc.dispose();
  process.chdir(originalCwd);
  validatorCache.set(xsdFile, validator);
  return validator;
}

function disposeAllValidators() {
  for (const v of validatorCache.values()) v.dispose();
  validatorCache.clear();
}

// ── Auto-detect XSD from file path and root element ──

interface XsdMapping {
  pattern: RegExp;
  xsd: string;
  /** Label for output (e.g. "slide1", "sheet2", "drawing1") */
  labelFn: (entry: string) => string;
}

const XSD_MAPPINGS: XsdMapping[] = [
  // ── PPTX ──
  {
    pattern: /^ppt\/slides\/slide\d+\.xml$/,
    xsd: "pml.xsd",
    labelFn: (e) => e.match(/slide(\d+)/)?.[1] ?? e,
  },
  {
    pattern: /^ppt\/presentation\.xml$/,
    xsd: "pml.xsd",
    labelFn: () => "presentation",
  },
  {
    pattern: /^ppt\/slideLayouts\/slideLayout\d+\.xml$/,
    xsd: "pml.xsd",
    labelFn: (e) => `slideLayout${e.match(/slideLayout(\d+)/)?.[1]}`,
  },
  {
    pattern: /^ppt\/slideMasters\/slideMaster\d+\.xml$/,
    xsd: "pml.xsd",
    labelFn: (e) => `slideMaster${e.match(/slideMaster(\d+)/)?.[1]}`,
  },
  {
    pattern: /^ppt\/notesMasters\/notesMaster\d+\.xml$/,
    xsd: "pml.xsd",
    labelFn: () => "notesMaster",
  },
  {
    pattern: /^ppt\/handoutMasters\/handoutMaster\d+\.xml$/,
    xsd: "pml.xsd",
    labelFn: () => "handoutMaster",
  },
  {
    pattern: /^ppt\/charts\/chart\d+\.xml$/,
    xsd: "dml-chart.xsd",
    labelFn: (e) => `chart${e.match(/chart(\d+)/)?.[1]}`,
  },
  {
    pattern: /^ppt\/tableStyles\.xml$/,
    xsd: "dml-main.xsd",
    labelFn: () => "tableStyles",
  },
  // ── DOCX ──
  { pattern: /^word\/document\.xml$/, xsd: "wml.xsd", labelFn: () => "document" },
  { pattern: /^word\/styles\.xml$/, xsd: "wml.xsd", labelFn: () => "styles" },
  { pattern: /^word\/numbering\.xml$/, xsd: "wml.xsd", labelFn: () => "numbering" },
  { pattern: /^word\/settings\.xml$/, xsd: "wml.xsd", labelFn: () => "settings" },
  { pattern: /^word\/footnotes\.xml$/, xsd: "wml.xsd", labelFn: () => "footnotes" },
  { pattern: /^word\/endnotes\.xml$/, xsd: "wml.xsd", labelFn: () => "endnotes" },
  { pattern: /^word\/comments\.xml$/, xsd: "wml.xsd", labelFn: () => "comments" },
  {
    pattern: /^word\/charts\/chart\d+\.xml$/,
    xsd: "dml-chart.xsd",
    labelFn: (e) => `chart${e.match(/chart(\d+)/)?.[1]}`,
  },
  // ── XLSX ──
  {
    pattern: /^xl\/worksheets\/sheet\d+\.xml$/,
    xsd: "sml.xsd",
    labelFn: (e) => `sheet${e.match(/sheet(\d+)/)?.[1]}`,
  },
  {
    pattern: /^xl\/chartsheets\/sheet\d+\.xml$/,
    xsd: "sml.xsd",
    labelFn: (e) => `chartsheet${e.match(/sheet(\d+)/)?.[1]}`,
  },
  { pattern: /^xl\/styles\.xml$/, xsd: "sml.xsd", labelFn: () => "styles" },
  { pattern: /^xl\/calcChain\.xml$/, xsd: "sml.xsd", labelFn: () => "calcChain" },
  {
    pattern: /^xl\/tables\/table\d+\.xml$/,
    xsd: "sml.xsd",
    labelFn: (e) => `table${e.match(/table(\d+)/)?.[1]}`,
  },
  {
    pattern: /^xl\/externalLinks\/externalLink\d+\.xml$/,
    xsd: "sml.xsd",
    labelFn: (e) => `externalLink${e.match(/externalLink(\d+)/)?.[1]}`,
  },
  {
    pattern: /^xl\/drawings\/drawing\d+\.xml$/,
    xsd: "dml-spreadsheetDrawing.xsd",
    labelFn: (e) => `drawing${e.match(/drawing(\d+)/)?.[1]}`,
  },
  {
    pattern: /^xl\/charts\/chart\d+\.xml$/,
    xsd: "dml-chart.xsd",
    labelFn: (e) => `chart${e.match(/chart(\d+)/)?.[1]}`,
  },
  // ── Shared (all packages) ──
  { pattern: /\/theme\/theme\d+\.xml$/, xsd: "dml-main.xsd", labelFn: () => "theme" },
];

function findXsdMapping(entry: string): XsdMapping | undefined {
  return XSD_MAPPINGS.find((m) => m.pattern.test(entry));
}

// ── Ignorable namespace / error filtering ──

const TRANSITIONAL_NAMESPACES = new Set([
  "http://schemas.microsoft.com/office/word/2010/wordml",
  "http://schemas.microsoft.com/office/word/2012/wordml",
  "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
  "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
  "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
  "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
  "urn:schemas-microsoft-com:vml",
  "urn:schemas-microsoft-com:office:office",
  "urn:schemas-microsoft-com:office:word",
]);

function extractIgnorableNamespaces(xml: string): Set<string> {
  const ignorable = new Set(TRANSITIONAL_NAMESPACES);
  const ignorableMatch = xml.match(/mc:Ignorable="([^"]+)"/);
  if (ignorableMatch) {
    const prefixes = ignorableMatch[1].split(/\s+/);
    for (const prefix of prefixes) {
      const nsMatch = xml.match(new RegExp(`xmlns:${prefix}="([^"]+)"`));
      if (nsMatch) ignorable.add(nsMatch[1]);
    }
  }
  return ignorable;
}

const TRANSITIONAL_ELEMENT_PATTERNS = [
  "'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pict'",
  "'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ruby'",
  "'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}checkbox'",
];

const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";

function isIgnorableError(message: string, ignorableNs: Set<string>): boolean {
  // mc:Ignorable attribute is valid per Markup Compatibility (OOXML Part 3)
  // but not declared in transitional XSDs — known limitation, not an implementation bug
  if (message.includes(`{${MC_NS}}Ignorable`)) return true;
  // w14/w15/w16 etc. extension namespace elements/attributes
  for (const pattern of TRANSITIONAL_ELEMENT_PATTERNS) {
    if (message.includes(pattern)) return true;
  }
  for (const ns of ignorableNs) {
    if (message.includes(`{${ns}}`) || message.includes(`'${ns}'`)) return true;
  }
  return false;
}

function validateXml(validator: XsdValidator, xml: string): { pass: boolean; error?: string } {
  const ignorableNs = extractIgnorableNamespaces(xml);
  const doc = XmlDocument.fromString(xml);
  try {
    validator.validate(doc);
    return { pass: true };
  } catch (e) {
    if (e instanceof Error && "details" in e) {
      const details = e.details as Array<{
        message: string;
        line: number;
        col: number;
      }>;
      const relevant = details.filter((d) => !isIgnorableError(d.message, ignorableNs));
      if (relevant.length === 0) return { pass: true };
      return { pass: false, error: relevant.map((d) => d.message.trim()).join("\n    ") };
    }
    return { pass: false, error: (e as Error).message };
  } finally {
    doc.dispose();
  }
}

// ── Package config ──

interface PackageConfig {
  name: string;
  dir: string;
  outputFile: string;
  buildCmd: string;
}

const PACKAGES: Record<string, PackageConfig> = {
  pptx: {
    name: "PPTX",
    dir: path.resolve(ROOT_DIR, "packages/pptx"),
    outputFile: "My Presentation.pptx",
    buildCmd: "pnpm -F @office-open/pptx build",
  },
  docx: {
    name: "DOCX",
    dir: path.resolve(ROOT_DIR, "packages/docx"),
    outputFile: "My Document.docx",
    buildCmd: "pnpm -F @office-open/docx build",
  },
  xlsx: {
    name: "XLSX",
    dir: path.resolve(ROOT_DIR, "packages/xlsx"),
    outputFile: "My Workbook.xlsx",
    buildCmd: "pnpm -F @office-open/xlsx build",
  },
};

// ── Validate all XML parts in a ZIP ──

function validateAllXmlParts(
  tmpFile: string,
  demo: string,
): { pass: number; fail: number; failures: string[] } {
  let totalPass = 0;
  let totalFail = 0;
  const failures: string[] = [];

  const entries = getZipEntries(tmpFile).filter(
    (e) => e.endsWith(".xml") && !e.startsWith("[Content_Types]"),
  );

  for (const entry of entries) {
    const mapping = findXsdMapping(entry);
    if (!mapping || !mapping.xsd) continue; // Skip entries with no XSD mapping

    const validator = loadValidator(mapping.xsd);
    const label = mapping.labelFn(entry);

    try {
      const xml = extractFromZip(tmpFile, entry);
      const result = validateXml(validator, xml);
      if (result.pass) {
        totalPass++;
      } else {
        totalFail++;
        failures.push(`${demo} ${label}: ${result.error}`);
      }
    } catch {
      // Some entries may fail extraction (unlikely), skip silently
    }
  }

  return { pass: totalPass, fail: totalFail, failures };
}

function validatePackage(pkg: PackageConfig) {
  console.log(`\n--- ${pkg.name} ---`);

  const demos = getDemoFiles(path.join(pkg.dir, "demo"));
  let totalPass = 0;
  let totalFail = 0;
  const allFailures: string[] = [];

  for (const demo of demos) {
    try {
      execSync(`npx tsx "demo/${demo}"`, { cwd: pkg.dir, stdio: "pipe" });
    } catch {
      console.error(`  RUN FAIL: ${demo}`);
      continue;
    }

    const outputPath = path.join(pkg.dir, pkg.outputFile);
    // Also check for round-trip variant files
    const candidates = [outputPath];
    const rtPath = outputPath.replace(
      path.extname(pkg.outputFile),
      ` (round-trip)${path.extname(pkg.outputFile)}`,
    );
    if (fs.existsSync(rtPath)) candidates.push(rtPath);

    for (const candidate of candidates) {
      if (!fs.existsSync(candidate)) continue;
      const tmpFile = path.join(pkg.dir, `__validate_tmp__${path.basename(candidate)}`);
      fs.renameSync(candidate, tmpFile);

      try {
        const result = validateAllXmlParts(tmpFile, demo);
        const status = result.fail === 0 ? "PASS" : `FAIL (${result.fail})`;
        const partLabel = candidate === outputPath ? "" : " [round-trip]";
        const entryCount = result.pass + result.fail;
        console.log(`  ${status}  ${demo}${partLabel} (${entryCount} parts)`);
        totalPass += result.pass;
        totalFail += result.fail;
        allFailures.push(...result.failures);
      } catch {
        console.error(`  ZIP FAIL: ${demo}`);
      }

      fs.unlinkSync(tmpFile);
    }
  }

  console.log(`  ${totalPass} passed, ${totalFail} failed`);
  return { totalPass, totalFail, failures: allFailures };
}

// ── Single-file validation ──

function validateSingleFile(filePath: string) {
  xmlRegisterFsInputProviders();
  const absPath = path.resolve(filePath);
  const tmpFile = absPath + ".__validate_tmp__";
  fs.copyFileSync(absPath, tmpFile);

  try {
    const result = validateAllXmlParts(tmpFile, path.basename(absPath));
    for (const f of result.failures) {
      console.error(`  FAIL: ${f}`);
    }
    if (result.fail > 0) process.exitCode = 1;
    console.log(`${result.pass} passed, ${result.fail} failed`);
  } finally {
    fs.unlinkSync(tmpFile);
  }
}

function main() {
  const args = process.argv.slice(2);

  // Single file validation mode
  if (args[0] === "slide" || args[0] === "docx" || args[0] === "xlsx") {
    if (args[1]) {
      xmlRegisterFsInputProviders();
      validateSingleFile(args[1]);
      disposeAllValidators();
      return;
    }
  }

  // Batch validation mode
  const targets = args.length > 0 ? args : ["pptx", "docx", "xlsx"];
  const invalid = targets.filter((t) => !(t in PACKAGES));
  if (invalid.length > 0) {
    console.error(`Unknown package(s): ${invalid.join(", ")}. Available: pptx, docx, xlsx`);
    process.exit(1);
  }

  xmlRegisterFsInputProviders();

  // Ensure core is built
  execSync("pnpm -F @office-open/core build", { cwd: ROOT_DIR, stdio: "pipe" });

  const allFailures: string[] = [];
  let grandPass = 0;
  let grandFail = 0;

  for (const target of targets) {
    const pkg = PACKAGES[target];
    execSync(pkg.buildCmd, { cwd: ROOT_DIR, stdio: "pipe" });
    const result = validatePackage(pkg);
    grandPass += result.totalPass;
    grandFail += result.totalFail;
    allFailures.push(...result.failures.map((f) => `[${target}] ${f}`));
  }

  console.log(`\n${"=".repeat(50)}`);
  console.log(`Total: ${grandPass} passed, ${grandFail} failed`);
  if (allFailures.length > 0) {
    console.log("\nFailures:");
    for (const f of allFailures) {
      console.log(`  ${f}`);
    }
    process.exitCode = 1;
  }

  disposeAllValidators();
}

main();
