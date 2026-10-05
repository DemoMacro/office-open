import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const PACKAGES = path.join(ROOT, "packages");
const CHECKED_PACKAGES = new Set(["odf", "odt", "ods", "odp"]);
const OOXML_PEERS: Record<string, string> = {
  odt: "@office-open/docx",
  ods: "@office-open/xlsx",
  odp: "@office-open/pptx",
};

interface PackageConfig {
  name: string;
  devDependencies: Set<string>;
  alwaysBundle: Set<string>;
}

function packageDirectories(): string[] {
  return fs
    .readdirSync(PACKAGES, { withFileTypes: true })
    .filter((entry) => entry.isDirectory() && CHECKED_PACKAGES.has(entry.name))
    .map((entry) => path.join(PACKAGES, entry.name));
}

function packageConfig(directory: string): PackageConfig {
  const manifestPath = path.join(directory, "package.json");
  const manifest = JSON.parse(fs.readFileSync(manifestPath, "utf8"));
  const configPath = path.join(directory, "vite.config.ts");
  const config = fs.readFileSync(configPath, "utf8");
  const alwaysBundleMatch = /alwaysBundle:\s*\[([^\]]+)\]/s.exec(config);
  const alwaysBundle = new Set(
    (alwaysBundleMatch?.[1] ?? "")
      .split(",")
      .map((value) => value.trim().replace(/^["']|["']$/g, ""))
      .filter(Boolean),
  );
  return {
    name: manifest.name,
    devDependencies: new Set(Object.keys(manifest.devDependencies ?? {})),
    alwaysBundle,
  };
}

function outputFiles(directory: string): string[] {
  const dist = path.join(directory, "dist");
  if (!fs.existsSync(dist)) return [];
  return fs.readdirSync(dist, { withFileTypes: true }).flatMap((entry) => {
    const full = path.join(dist, entry.name);
    if (entry.isDirectory()) return outputFiles(full);
    return entry.name.endsWith(".mjs") ? [full] : [];
  });
}

function importedPackages(source: string): Set<string> {
  return new Set(
    [...source.matchAll(/(?:from\s*|import\s*|import\()\s*["']([^"']+)["']/g)]
      .map((match) => match[1]!)
      .filter(
        (specifier) => specifier === "@office-open/xml" || specifier.startsWith("@office-open/"),
      )
      .map((specifier) => specifier.split("/").slice(0, 2).join("/")),
  );
}

const problems: string[] = [];

for (const directory of packageDirectories()) {
  const config = packageConfig(directory);
  const files = outputFiles(directory);
  if (files.length === 0) {
    problems.push(`${config.name}: dist output is missing`);
    continue;
  }

  for (const file of files) {
    const source = fs.readFileSync(file, "utf8");
    const imports = importedPackages(source);
    const relativeFile = path.relative(ROOT, file).replaceAll("\\", "/");
    const isDeclaration = file.endsWith(".d.mts");

    for (const imported of imports) {
      if (!isDeclaration && config.devDependencies.has(imported)) {
        problems.push(`${relativeFile}: runtime JS imports devDependency ${imported}`);
      }
      if (isDeclaration && config.alwaysBundle.has(imported)) {
        problems.push(`${relativeFile}: declaration retains inlined dependency ${imported}`);
      }
      if (!isDeclaration && imported === OOXML_PEERS[path.basename(directory)]) {
        problems.push(`${relativeFile}: runtime imports OOXML peer ${imported}`);
      }
    }
  }
}

if (problems.length > 0) {
  console.error(`package outputs: ${problems.length} problem(s)`);
  for (const problem of problems) console.error(`  - ${problem}`);
  process.exit(1);
}

console.log(`package outputs: OK (${CHECKED_PACKAGES.size} packages)`);
