/**
 * Extract Options-interface field names from a TypeScript interface definition.
 *
 * Test-only utility (ts-morph is a root devDependency). Lives under core/tests/
 * so it is isolated from src/ — it never enters the package barrel or the
 * published bundle, which would otherwise pull in the TypeScript compiler.
 *
 * Reads field names via the type checker so `&` intersection types resolve to
 * their full merged member set: a declaration like
 * `ParagraphPropertiesOptions = {...} & ParagraphStylePropertiesOptions`
 * would lose inherited members via InterfaceDeclaration.getProperties(), but
 * Type.getProperties() returns the union. Index signatures are excluded
 * automatically (they are not "properties" to the checker).
 *
 * @module
 */
import fs from "node:fs";
import path from "node:path";

import { Project, type SourceFile } from "ts-morph";

let project: Project | undefined;
let root: string | undefined;

/** Walk up from cwd to the workspace root (package.json name "@office-open/monorepo"). */
function monorepoRoot(): string {
  if (root) return root;
  let dir = process.cwd();
  for (let i = 0; i < 8; i++) {
    try {
      const name = JSON.parse(fs.readFileSync(path.join(dir, "package.json"), "utf8")).name;
      if (name === "@office-open/monorepo") return (root = dir);
    } catch {
      // package.json unreadable — keep walking up
    }
    const parent = path.dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }
  return (root = process.cwd());
}

function tsProject(): Project {
  if (!project) {
    project = new Project({
      compilerOptions: { declaration: false, noEmit: true },
      skipAddingFilesFromTsConfig: true,
    });
  }
  return project;
}

function loadSourceFile(sourceFilePath: string): SourceFile {
  const abs = path.isAbsolute(sourceFilePath)
    ? sourceFilePath
    : path.resolve(monorepoRoot(), sourceFilePath);
  const p = tsProject();
  return p.getSourceFile(abs) ?? p.addSourceFileAtPath(abs);
}

/**
 * Read the field names of an interface. Resolves `&` intersection and
 * inheritance, so a composed options interface returns its full field set.
 * Returns the names sorted alphabetically for stable comparison.
 */
export function extractInterfaceFields(interfaceName: string, sourceFilePath: string): string[] {
  const sf = loadSourceFile(sourceFilePath);
  // Options types may be declared as `interface` (CorePropertiesInput) or as a
  // `type` alias over an intersection (ParagraphPropertiesOptions = {...} & ...).
  // Both expose getType(), which resolves the intersection to its merged members.
  const decl = sf.getInterface(interfaceName) ?? sf.getTypeAlias(interfaceName);
  if (!decl) {
    throw new Error(`extractInterfaceFields: "${interfaceName}" not found in ${sourceFilePath}`);
  }
  return decl
    .getType()
    .getProperties()
    .map((s) => s.getName())
    .sort();
}
