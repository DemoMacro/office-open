/**
 * en/zh documentation parity gate.
 *
 * The two language trees under docs/content are translations of each other:
 * same file set, same section structure, same example blocks (Astro Docs'
 * i18n model — "faithful representations without changes to structure").
 * This gate enforces the structural half so a translation cannot silently
 * drift behind its English source:
 *
 *   1. file sets are 1:1 (both directions);
 *   2. per file pair: code-block count, heading count, and ::api-example
 *      directive count match;
 *   3. internal section links follow the site's route convention — no
 *      numeric ordering prefix and no .md suffix (`./export` resolves to
 *      `18.export.md` because Nuxt Content strips `NN.` prefixes), and the
 *      target section must exist.
 *
 * Usage:
 *   npx tsx scripts/docs-parity-check.ts
 */
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const CONTENT = path.join(ROOT, "docs", "content");
const EN = path.join(CONTENT, "en");
const ZH = path.join(CONTENT, "zh");

function walk(dir: string): string[] {
  const out: string[] = [];
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) out.push(...walk(full));
    else if (entry.name.endsWith(".md")) out.push(full);
  }
  return out;
}

/** Count fenced code blocks, headings, and ::api-example directives. */
function profile(text: string): { blocks: number; headings: number; examples: number } {
  const lines = text.split("\n");
  let blocks = 0;
  let headings = 0;
  let examples = 0;
  let inFence = false;
  for (const line of lines) {
    if (line.trimStart().startsWith("```")) {
      if (!inFence) blocks++;
      inFence = !inFence;
      continue;
    }
    if (inFence) continue;
    if (/^#{1,6} /.test(line)) headings++;
    if (line.includes("::api-example")) examples++;
  }
  return { blocks, headings, examples };
}

interface LinkIssue {
  file: string;
  target: string;
  reason: string;
}

/** Validate internal `](./...)` links against the route convention. */
function checkLinks(relFile: string, text: string): LinkIssue[] {
  const issues: LinkIssue[] = [];
  const dir = path.join(CONTENT, path.dirname(relFile));
  for (const match of text.matchAll(/\]\((\.[^)#\s]*)(#[^)\s]*)?\)/g)) {
    const target = match[1]!;
    const clean = target.replace(/^\.\//, "");
    if (/^\d+\./.test(clean)) {
      issues.push({ file: relFile, target, reason: "numeric ordering prefix (route strips NN.)" });
      continue;
    }
    if (clean.endsWith(".md")) {
      issues.push({ file: relFile, target, reason: ".md suffix (routes have none)" });
      continue;
    }
    if (!fs.existsSync(path.join(dir, clean))) continue; // external-ish or asset, skip
    // resolve `./slug` against sibling files whose route slug matches
    const siblings = fs.readdirSync(dir);
    if (!siblings.some((s) => s.replace(/^\d+\./, "").replace(/\.md$/, "") === clean)) {
      issues.push({ file: relFile, target, reason: "no sibling section routes to this slug" });
    }
  }
  return issues;
}

const enFiles = walk(EN).map((f) => path.relative(EN, f));
const zhFiles = walk(ZH).map((f) => path.relative(ZH, f));

const problems: string[] = [];

for (const f of enFiles) {
  if (!zhFiles.includes(f)) problems.push(`missing zh counterpart: en/${f}`);
}
for (const f of zhFiles) {
  if (!enFiles.includes(f)) problems.push(`zh file with no en source: zh/${f}`);
}

let linkIssues = 0;
for (const f of enFiles) {
  const enText = fs.readFileSync(path.join(EN, f), "utf8");
  const zhPath = path.join(ZH, f);
  if (!fs.existsSync(zhPath)) continue;
  const zhText = fs.readFileSync(zhPath, "utf8");
  const pe = profile(enText);
  const pz = profile(zhText);
  if (pe.blocks !== pz.blocks || pe.headings !== pz.headings || pe.examples !== pz.examples) {
    problems.push(
      `${f}: en/zh structure drift — ` +
        `blocks ${pe.blocks} vs ${pz.blocks}, headings ${pe.headings} vs ${pz.headings}, ` +
        `api-examples ${pe.examples} vs ${pz.examples}`,
    );
  }
  for (const issue of [...checkLinks(`en/${f}`, enText), ...checkLinks(`zh/${f}`, zhText)]) {
    problems.push(`${issue.file}: link "${issue.target}" — ${issue.reason}`);
    linkIssues++;
  }
}

if (problems.length > 0) {
  console.error(`docs parity: ${problems.length} problem(s)`);
  for (const p of problems) console.error(`  - ${p}`);
  process.exit(1);
}
console.log(`docs parity: OK (${enFiles.length} file pairs, ${linkIssues} link issues)`);
