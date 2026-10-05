/**
 * Enforce the workspace dependency graph in addition to declared dependencies.
 *
 * Target-format packages that are not present in an integration checkout are
 * reported as skipped: a missing directory must not turn into a false failure.
 */
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const PACKAGES = path.join(ROOT, "packages");
const WORKSPACE_PREFIX = "@office-open/";

interface ArchitectureRule {
  runtime?: ReadonlySet<string>;
  type?: ReadonlySet<string>;
  denied?: ReadonlySet<string>;
}

const FORMAT_PACKAGES = new Set([
  "doc",
  "docx",
  "xls",
  "xlsx",
  "ppt",
  "pptx",
  "rtf",
  "odf",
  "odt",
  "ods",
  "odp",
]);

const ARCHITECTURE: Record<string, ArchitectureRule> = {
  core: { denied: FORMAT_PACKAGES },
  odf: {
    runtime: new Set(["xml"]),
    type: new Set(["core"]),
    denied: new Set(["docx", "xlsx", "pptx", "odt", "ods", "odp"]),
  },
  odt: {
    runtime: new Set(["core", "odf"]),
    type: new Set(["core", "xml", "docx"]),
  },
  ods: {
    runtime: new Set(["odf"]),
    type: new Set(["core", "xml", "xlsx"]),
  },
  odp: {
    runtime: new Set(["core", "odf"]),
    type: new Set(["core", "xml", "pptx"]),
  },
};

const TARGET_PACKAGES = new Set(["odf", "odt", "ods", "odp"]);

function stripComments(source: string): string {
  return source
    .replace(/\/\*[\s\S]*?\*\//g, (match) => match.replace(/[^\n]/g, " "))
    .replace(/\/\/[^\n]*/g, "");
}

function packageName(moduleSpecifier: string): string {
  const packagePath = moduleSpecifier.slice(WORKSPACE_PREFIX.length);
  return `${WORKSPACE_PREFIX}${packagePath.split("/")[0]}`;
}

function importedPackages(source: string): Array<{ name: string; typeOnly: boolean }> {
  const code = stripComments(source);
  const imports = new Map<string, { name: string; typeOnly: boolean }>();

  for (const match of code.matchAll(
    /(?:^|\n)\s*import\s+(type\s+)?[\w*$\s{},]+?\s+from\s+["'](@office-open\/[^"']+)["']/g,
  )) {
    const name = packageName(match[2]!);
    imports.set(name, { name, typeOnly: Boolean(match[1]) });
  }

  for (const match of code.matchAll(/(?:^|\n)\s*import\s+["'](@office-open\/[^"']+)["']/g)) {
    const name = packageName(match[1]!);
    imports.set(name, { name, typeOnly: false });
  }

  for (const match of code.matchAll(/import\(\s*["'](@office-open\/[^"']+)["']\s*\)/g)) {
    const name = packageName(match[1]!);
    imports.set(name, { name, typeOnly: false });
  }

  for (const match of code.matchAll(
    /(?:^|\n)\s*export\s+(type\s+)?(?:\*(?:\s+as\s+[\w$]+)?|\{[^}]*\})\s+from\s+["'](@office-open\/[^"']+)["']/g,
  )) {
    const name = packageName(match[2]!);
    imports.set(name, { name, typeOnly: Boolean(match[1]) });
  }

  return [...imports.values()];
}

function exportedStars(source: string): string[] {
  const code = stripComments(source);
  return [
    ...code.matchAll(
      /(?:^|\n)\s*export\s+(?:type\s+)?\*(?:\s+as\s+[\w$]+)?\s+from\s+["'](@office-open\/[^"']+)["']/g,
    ),
  ].map((match) => packageName(match[1]!));
}

function declaredDependencies(packageDirectory: string): Set<string> {
  const manifest = JSON.parse(fs.readFileSync(path.join(packageDirectory, "package.json"), "utf8"));
  return new Set(
    ["dependencies", "devDependencies", "peerDependencies", "optionalDependencies"].flatMap(
      (field) => Object.keys(manifest[field] ?? {}),
    ),
  );
}

function packageDirectory(name: string): string {
  return path.join(PACKAGES, name.slice(WORKSPACE_PREFIX.length));
}

function packageExists(name: string): boolean {
  return fs.existsSync(path.join(packageDirectory(name), "package.json"));
}

function walkSources(directory: string): string[] {
  return fs.readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const full = path.join(directory, entry.name);
    if (entry.isDirectory()) return walkSources(full);
    return entry.name.endsWith(".ts") && !entry.name.endsWith(".d.ts") ? [full] : [];
  });
}

const problems: string[] = [];
const skipped: string[] = [];

for (const name of TARGET_PACKAGES) {
  if (!packageExists(`${WORKSPACE_PREFIX}${name}`)) skipped.push(name);
}
for (const entry of fs.readdirSync(PACKAGES, { withFileTypes: true })) {
  if (!entry.isDirectory() || entry.name === "node_modules") continue;
  const directory = path.join(PACKAGES, entry.name);
  const name = `${WORKSPACE_PREFIX}${entry.name}`;
  if (!fs.existsSync(path.join(directory, "package.json"))) continue;
  if (TARGET_PACKAGES.has(entry.name) && !packageExists(name)) continue;

  const declared = declaredDependencies(directory);
  for (const file of walkSources(directory)) {
    const relativeFile = path.relative(ROOT, file).replaceAll("\\", "/");
    const source = fs.readFileSync(file, "utf8");
    const isPackageTest = relativeFile.endsWith(".spec.ts");

    for (const imported of importedPackages(source)) {
      if (imported.name === name) continue;
      if (!declared.has(imported.name)) {
        problems.push(`${relativeFile} imports undeclared ${imported.name}`);
        continue;
      }

      const rule = ARCHITECTURE[entry.name];
      if (!rule) continue;
      const importedId = imported.name.slice(WORKSPACE_PREFIX.length);
      if (rule.denied?.has(importedId)) {
        problems.push(`${relativeFile} imports denied ${imported.name}`);
        continue;
      }
      if (rule.runtime || rule.type) {
        const allowed = imported.typeOnly
          ? new Set([...(rule.runtime ?? []), ...(rule.type ?? [])])
          : (rule.runtime ?? new Set());
        if (!allowed.has(importedId)) {
          problems.push(
            `${relativeFile} ${imported.typeOnly ? "type" : "runtime"} import violates ${entry.name} topology: ${imported.name}`,
          );
        }
      }
    }

    for (const exported of exportedStars(source)) {
      if (entry.name === "office-open") continue;
      problems.push(`${relativeFile} re-exports * across package boundary: ${exported}`);
    }
  }
}

if (skipped.length > 0) {
  console.warn(`package boundaries: ${skipped.length} target package(s) skipped`);
  for (const name of skipped) console.warn(`  - ${name}: directory not integrated`);
}

if (problems.length > 0) {
  console.error(`package boundaries: ${problems.length} problem(s)`);
  for (const problem of problems) console.error(`  - ${problem}`);
  process.exit(1);
}

console.log(`package boundaries: OK${skipped.length > 0 ? ` (${skipped.length} skipped)` : ""}`);
