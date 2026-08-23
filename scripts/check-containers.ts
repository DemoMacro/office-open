/**
 * Container-model gate: validate the XML inside generated OOXML packages
 * against scripts/container-models.json — child legality, slot order, and
 * single-round pick exclusivity. A fast local gate over the demo artifacts
 * (no .NET, no libxml2); the extracted XSD content-model table is the golden
 * source (`pnpm models:generate`).
 *
 * Usage:
 *   npx tsx scripts/check-containers.ts            # all demo artifacts (packages/<pkg>/.temp)
 *   npx tsx scripts/check-containers.ts <file>...   # specific package files
 *
 * Exits non-zero when any violation is found. mc:AlternateContent subtrees
 * are skipped (MCE version-negotiated content); non-ISO prefixes (wp14/x14…)
 * pass through — MCE adjudicates those, not this gate.
 */
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

import { unzipSync } from "fflate";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT_DIR = path.resolve(__dirname, "..");
const MODELS_PATH = path.join(ROOT_DIR, "scripts", "container-models.json");

interface Slot {
  elements?: string[];
  min?: number;
  max?: number;
  pick?: "one" | "any";
  any?: string[];
}
interface ContainerModel {
  slots: Slot[];
  variantElements?: string[];
}
interface ModelsFile {
  containers: Record<string, ContainerModel>;
  namespaces: Record<string, string>;
}

const models = JSON.parse(fs.readFileSync(MODELS_PATH, "utf8")) as ModelsFile;
const containers = models.containers;
// any-wildcard namespaces expressed as prefixes (models.namespaces maps prefix→URI)
const nsUriToPrefix = new Map(Object.entries(models.namespaces).map(([p, u]) => [u, p]));

interface Prepared {
  members: Set<string>;
  slotOf: Map<string, number>;
  multiSlot: Set<string>;
  exclusiveSlot: Map<string, number>; // member → pick-one max-1 slot (-1 = ambiguous)
  anyPrefixes: string[];
  anyAll: boolean;
  otherExcluded: string | null;
}

// per container type (lazily built): member set, slot index per member,
// single-round exclusive members (pick one, max 1)
const prepared = new Map<string, Prepared | null>();
function prepare(qname: string): Prepared | null {
  let p = prepared.get(qname);
  if (p !== undefined) return p;
  const c = containers[qname];
  if (!c) {
    prepared.set(qname, null);
    return null;
  }
  const members = new Set<string>();
  const slotOf = new Map<string, number>();
  const multiSlot = new Set<string>();
  const exclusiveSlot = new Map<string, number>(); // exclusivity is per slot
  const anyPrefixes: string[] = [];
  let anyAll = false;
  let otherExcluded: string | null = null; // prefix to exclude for ##other
  for (let i = 0; i < c.slots.length; i++) {
    const s = c.slots[i]!;
    if (s.any) {
      for (const uri of s.any) {
        if (uri === "##any") anyAll = true;
        else if (uri === "##other")
          otherExcluded = qname.includes(":") ? qname.slice(0, qname.indexOf(":")) : "";
        else {
          const px = nsUriToPrefix.get(uri);
          if (px !== undefined) anyPrefixes.push(px);
        }
      }
      continue;
    }
    if (s.pick === "one" && s.max === 1) {
      for (const e of s.elements ?? []) {
        // an element in several pick-one slots is only safely exclusive when
        // every occurrence is the same slot — skip ambiguity
        if (exclusiveSlot.has(e)) exclusiveSlot.set(e, -1);
        else exclusiveSlot.set(e, i);
      }
    }
    for (const e of s.elements ?? []) {
      members.add(e);
      if (slotOf.has(e)) multiSlot.add(e);
      else slotOf.set(e, i);
    }
  }
  // context-dependent elements (w:sdtContent block/run/cell/row variants)
  // cannot be checked against one picked model — the table carries the union
  if (c.variantElements) for (const e of c.variantElements) members.add(e);
  p = { members, slotOf, multiSlot, exclusiveSlot, anyPrefixes, anyAll, otherExcluded };
  prepared.set(qname, p);
  return p;
}

const violations: string[] = [];

// ISO-transitional table only knows its own namespaces; other prefixes are
// vendor extensions (wp14/x14/…) adjudicated by MCE, not by this gate
const knownPrefixes = new Set(Object.keys(models.namespaces));

interface StackEntry {
  qname: string;
  children: string[];
}

function checkContainer(qname: string, children: string[], at: string): void {
  const p = prepare(qname);
  if (!p) return;
  // child set legality
  for (const child of children) {
    if (p.members.has(child)) continue;
    if (child.startsWith("mc:")) continue;
    const px = child.includes(":") ? child.slice(0, child.indexOf(":")) : "";
    if (!knownPrefixes.has(px)) continue;
    if (p.anyAll) continue;
    if (p.otherExcluded !== null && px !== p.otherExcluded) continue;
    if (p.anyPrefixes.includes(px) || p.anyPrefixes.includes("")) continue;
    violations.push(`${at}: <${qname}> child ${child} not in content model`);
  }
  // slot order: slot index sequence over modeled children must be non-decreasing
  let prev = -1;
  let prevChild = "";
  for (const child of children) {
    if (p.multiSlot.has(child) || !p.slotOf.has(child)) continue;
    const idx = p.slotOf.get(child)!;
    if (idx < prev) {
      violations.push(
        `${at}: <${qname}> order: ${child} (slot ${idx}) after ${prevChild} (slot ${prev})`,
      );
    }
    prev = idx;
    prevChild = child;
  }
  // single-round exclusivity: members of the same pick-one max-1 slot co-occurring
  const perSlot = new Map<number, Set<string>>();
  for (const child of children) {
    const idx = p.exclusiveSlot.get(child);
    if (idx === undefined || idx < 0) continue;
    if (!perSlot.has(idx)) perSlot.set(idx, new Set());
    perSlot.get(idx)!.add(child);
  }
  for (const names of perSlot.values()) {
    if (names.size > 1) {
      violations.push(`${at}: <${qname}> exclusive members co-occur: ${[...names].join(" + ")}`);
    }
  }
}

function parentPath(stack: StackEntry[]): string {
  const tail = stack.slice(-2);
  return tail.length ? tail.map((s) => s.qname).join(">") : (stack.at(-1)?.qname ?? "?");
}

function checkFile(file: string): void {
  const zip = unzipSync(new Uint8Array(fs.readFileSync(file)));
  const dec = new TextDecoder();

  for (const name of Object.keys(zip)) {
    if (!name.endsWith(".xml") || name.endsWith(".rels") || name === "[Content_Types].xml")
      continue;
    const xml = dec.decode(zip[name]!);
    if (!/<(w|p|a|c|dgm|pic|wp|xdr|lc|cdr|m|v|o|w10|x|pvml):/.test(xml)) continue;

    // stack entries: { qname, children }
    const stack: StackEntry[] = [];
    let mcDepth = 0; // inside mc:AlternateContent the content is version-negotiated — not validated
    const tagRe = /<(\/?)([A-Za-z0-9]+:)?([A-Za-z0-9]+)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>/g;
    let m: RegExpExecArray | null;
    let offset = 0;
    while ((m = tagRe.exec(xml)) !== null) {
      const closing = m[1] === "/";
      const prefix = m[2] ?? "";
      const qname = `${prefix}${m[3]}`;
      const selfClose = m[5] === "/";
      if (closing) {
        const top = stack.pop();
        if (qname === "mc:AlternateContent") mcDepth--;
        if (top && top.qname === qname && mcDepth === 0) {
          const parent = stack.at(-1);
          if (parent) {
            parent.children.push(qname);
            checkContainer(parent.qname, parent.children, `${name} > ${parentPath(stack)}`);
          }
        }
        continue;
      }
      if (qname === "mc:AlternateContent") mcDepth++;
      if (selfClose) {
        const parent = stack.at(-1);
        if (parent && mcDepth === 0) parent.children.push(qname);
        continue;
      }
      stack.push({ qname, children: [] });
      offset++;
      if (offset > 2e6) break; // safety
    }
  }
}

// ── Main ──

const args = process.argv.slice(2);
let files: string[];
if (args.length > 0) {
  files = args;
} else {
  // default: every demo artifact the validate gate produces
  files = [];
  for (const pkg of ["docx", "pptx", "xlsx"]) {
    const dir = path.join(ROOT_DIR, "packages", pkg, ".temp");
    if (!fs.existsSync(dir)) continue;
    for (const name of fs.readdirSync(dir)) {
      if (/\.(docx|pptx|xlsx)$/.test(name)) files.push(path.join(dir, name));
    }
  }
  if (files.length === 0) {
    console.error(
      "no demo artifacts found — run `pnpm validate` first (writes packages/<pkg>/.temp/)",
    );
    process.exit(1);
  }
}

for (const file of files) checkFile(file);

// ── report ──
const byKind = new Map<string, number>();
for (const v of violations) {
  const kind = v.includes("not in content model")
    ? "illegal-child"
    : v.includes("order:")
      ? "order"
      : "exclusive";
  byKind.set(kind, (byKind.get(kind) ?? 0) + 1);
}
console.log(`${files.length} file(s): ${violations.length} violations`, Object.fromEntries(byKind));
const shown = 40;
for (const v of violations.slice(0, shown)) console.log("  " + v);
if (violations.length > shown) console.log(`  … ${violations.length - shown} more`);
if (violations.length > 0) process.exit(1);
