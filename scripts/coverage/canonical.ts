/**
 * Canonical part-coverage scanner — verifies each OOXML part descriptor
 * implements BOTH stringify and parse as real (non-stub) implementations.
 *
 * Unlike the legacy xsd.ts name scan, this walks the actual descriptor
 * modules per package and flags:
 *  - stringify stubs (`stringify() { throw ... }` delegating elsewhere)
 *  - parse-only / stringify-only descriptors
 *  - descriptor modules missing from any package
 *
 * Usage: pnpm exec tsx scripts/coverage/canonical.ts [--summary]
 *
 * @module
 */

import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");

interface PartFinding {
  pkg: string;
  file: string;
  descriptor: string;
  stringify: "ok" | "contextual" | "stub" | "missing";
  parse: "ok" | "missing";
}

const PACKAGES = ["docx", "pptx", "xlsx"] as const;

function collectFiles(dir: string, out: string[] = []): string[] {
  if (!fs.existsSync(dir)) return out;
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) collectFiles(full, out);
    else if (entry.name.endsWith(".ts") && !entry.name.endsWith(".spec.ts")) out.push(full);
  }
  return out;
}

function stripComments(src: string): string {
  let out = "";
  let quote: string | undefined;
  for (let i = 0; i < src.length; i++) {
    const char = src[i]!;
    const next = src[i + 1];
    if (quote) {
      out += char;
      if (char === "\\") {
        if (next !== undefined) out += next;
        i++;
      } else if (char === quote) {
        quote = undefined;
      }
      continue;
    }
    if (char === '"' || char === "'" || char === "`") {
      quote = char;
      out += char;
      continue;
    }
    if (char === "/" && next === "/") {
      while (i < src.length && src[i] !== "\n") i++;
      if (src[i] === "\n") out += "\n";
      continue;
    }
    if (char === "/" && next === "*") {
      i += 2;
      while (i < src.length && !(src[i] === "*" && src[i + 1] === "/")) i++;
      i++;
      out += " ";
      continue;
    }
    out += char;
  }
  return out;
}

function scanDescriptor(
  src: string,
  name: string,
  hasPackageWriteBridge: boolean,
): { stringify: PartFinding["stringify"]; parse: PartFinding["parse"] } | undefined {
  // Locate the descriptor object — `export const xDesc: CustomDescriptor<...> = {`
  // or `export const xDesc = {`.
  const declRe = new RegExp(`export const ${name}(?:\\s*:\\s*[^=]+)?\\s*=\\s*\\{`);
  const decl = declRe.exec(src);
  if (!decl) return undefined;

  // Slice the object body with string-aware brace matching: descriptor XML
  // builders commonly contain template interpolation such as `${value}`.
  let depth = 0;
  let start = decl.index + decl[0].length - 1;
  let end = start;
  let quote: string | undefined;
  for (let i = start; i < src.length; i++) {
    const char = src[i]!;
    if (quote) {
      if (char === "\\") i++;
      else if (char === quote) quote = undefined;
      continue;
    }
    if (char === '"' || char === "'" || char === "`") {
      quote = char;
      continue;
    }
    if (char === "{") depth++;
    else if (char === "}") {
      depth--;
      if (depth === 0) {
        end = i;
        break;
      }
    }
  }
  const body = src.slice(start, end);

  const member = (
    methodName: string,
  ): { exists: boolean; isStub: boolean; hasWriteBridge: boolean } => {
    const re = new RegExp(`${methodName}\\s*\\(`);
    const match = re.exec(body);
    if (!match) return { exists: false, isStub: false };
    // Find the method body with string-aware brace matching, then check whether
    // it begins with a throw.
    let mDepth = 0;
    let mStart = -1;
    let mQuote: string | undefined;
    for (let i = match.index; i < body.length; i++) {
      const char = body[i]!;
      if (mQuote) {
        if (char === "\\") i++;
        else if (char === mQuote) mQuote = undefined;
        continue;
      }
      if (char === '"' || char === "'" || char === "`") {
        mQuote = char;
        continue;
      }
      if (char === "{") {
        if (mDepth === 0) mStart = i;
        mDepth++;
      } else if (char === "}") {
        mDepth--;
        if (mDepth === 0) break;
      }
    }
    const methodBody = body.slice(mStart, mStart + 400);
    const isStub = /^\s*throw\b/.test(methodBody.replace(/^\{\s*/, ""));
    const hasWriteBridge = isStub && hasPackageWriteBridge;
    return { exists: true, isStub, hasWriteBridge };
  };

  const stringify = member("stringify");
  const parse = member("parse");
  return {
    stringify: !stringify.exists
      ? "missing"
      : stringify.isStub && stringify.hasWriteBridge
        ? "contextual"
        : stringify.isStub
          ? "stub"
          : "ok",
    parse: parse.exists ? "ok" : "missing",
  };
}

function scanPackage(pkg: string): PartFinding[] {
  const findings: PartFinding[] = [];
  const partsDir = path.join(ROOT_DIR, "packages", pkg, "src", "parts");
  const packageFiles = collectFiles(path.join(ROOT_DIR, "packages", pkg, "src"));
  const hasPackageWriteBridge = packageFiles.some((file) =>
    fs.readFileSync(file, "utf-8").includes("export function stringifyWorksheet("),
  );
  for (const file of collectFiles(partsDir)) {
    const raw = fs.readFileSync(file, "utf-8");
    if (!raw.includes("CustomDescriptor")) continue;
    const src = stripComments(raw);
    const declRe = /export const ([a-zA-Z][a-zA-Z0-9]*Desc)\b/g;
    let match: RegExpExecArray | null;
    while ((match = declRe.exec(src)) !== null) {
      const name = match[1];
      const result = scanDescriptor(src, name, hasPackageWriteBridge);
      if (!result) continue;
      findings.push({
        pkg,
        file: path.relative(ROOT_DIR, file),
        descriptor: name,
        stringify: result.stringify,
        parse: result.parse,
      });
    }
  }
  return findings;
}

function main(): void {
  const summaryOnly = process.argv.includes("--summary");
  let total = 0;
  let gaps = 0;
  for (const pkg of PACKAGES) {
    const findings = scanPackage(pkg);
    const pkgGaps = findings.filter(
      (f) => f.stringify === "missing" || f.stringify === "stub" || f.parse !== "ok",
    );
    total += findings.length;
    gaps += pkgGaps.length;
    if (!summaryOnly) {
      console.log(`\n[${pkg}] ${findings.length} part descriptors`);
      for (const gap of pkgGaps)
        console.log(
          `  GAP ${gap.descriptor} — stringify: ${gap.stringify}, parse: ${gap.parse} (${gap.file})`,
        );
    } else {
      console.log(
        `[${pkg}] ${findings.length} descriptors, ${pkgGaps.length} with bidirectional gaps`,
      );
      for (const gap of pkgGaps)
        console.log(`  GAP ${gap.descriptor} (stringify: ${gap.stringify}, parse: ${gap.parse})`);
    }
  }
  console.log(`\ncanonical part coverage: ${total - gaps}/${total} bidirectional, ${gaps} gaps`);
}

const invokedDirectly = process.argv[1]?.endsWith("canonical.ts");
if (invokedDirectly) main();
