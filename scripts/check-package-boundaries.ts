/**
 * Verify that every workspace import in a package source tree is declared by
 * that package. This catches hoisted "ghost dependencies" before they break
 * isolated installs, without yet enforcing the target package topology.
 */
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const PACKAGES = path.join(ROOT, "packages");
const WORKSPACE_IMPORT =
  /(?:^|\n)\s*(?:import|export)(?:\s+type)?\s[^;]*?from\s+["'](@office-open\/[^"']+)(?:\/[^"']*)?["']|import\(["'](@office-open\/[^"']+)["']\)/g;

function walkSources(directory: string): string[] {
  return fs.readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const full = path.join(directory, entry.name);
    if (entry.isDirectory()) return walkSources(full);
    return entry.name.endsWith(".ts") && !entry.name.endsWith(".d.ts") ? [full] : [];
  });
}

function declaredDependencies(packageDirectory: string): Set<string> {
  const manifest = JSON.parse(fs.readFileSync(path.join(packageDirectory, "package.json"), "utf8"));
  return new Set(
    ["dependencies", "devDependencies", "peerDependencies", "optionalDependencies"].flatMap(
      (field) => Object.keys(manifest[field] ?? {}),
    ),
  );
}

function importedPackages(source: string): Set<string> {
  return new Set(
    [...source.matchAll(WORKSPACE_IMPORT)]
      .map((match) => match[1] ?? match[2])
      .filter((name): name is string => Boolean(name))
      .map((name) => {
        const [scope, packageName] = name.split("/");
        return `${scope}/${packageName!.split("/")[0]}`;
      }),
  );
}

const problems: string[] = [];
for (const entry of fs.readdirSync(PACKAGES, { withFileTypes: true })) {
  if (!entry.isDirectory() || entry.name === "node_modules") continue;
  const directory = path.join(PACKAGES, entry.name);
  if (!fs.existsSync(path.join(directory, "package.json"))) continue;
  const declared = declaredDependencies(directory);
  for (const file of walkSources(directory)) {
    const imports = importedPackages(fs.readFileSync(file, "utf8"));
    for (const name of imports) {
      if (name === `@office-open/${entry.name}`) continue;
      if (!declared.has(name)) {
        problems.push(`${path.relative(ROOT, file)} imports undeclared ${name}`);
      }
    }
  }
}

if (problems.length > 0) {
  console.error(`package boundaries: ${problems.length} problem(s)`);
  for (const problem of problems) console.error(`  - ${problem}`);
  process.exit(1);
}
console.log("package boundaries: OK");
