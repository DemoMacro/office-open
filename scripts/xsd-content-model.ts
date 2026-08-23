/**
 * XSD content-model extractor
 *
 * Parses the OOXML transitional XSDs, flattens xsd:group reference chains, and
 * emits per-element ordered child-slot tables to scripts/container-models.json.
 * That table is the golden source for container gates: the full-sample
 * generator narrows emitted combinations against it, and descriptor lint
 * checks emit order/parse coverage against it.
 *
 * Slot semantics:
 * - slots are ordered where the XSD says sequence; unordered pick groups where
 *   the XSD says choice
 * - pick "one": members are mutually exclusive per round (choice, bounded
 *   repeat), pick "any": members may appear in any order any number of times
 *   (choice under an unbounded reference)
 * - xsd:any wildcard slots keep their namespace list for pass-through domains
 *
 * Usage:
 *   pnpm tsx scripts/xsd-content-model.ts            # write container-models.json
 *   pnpm tsx scripts/xsd-content-model.ts --check    # diff against the committed file
 */
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SCHEMA_DIR = path.resolve(__dirname, "../ooxml-schemas/transitional");
const OUT_FILE = path.resolve(__dirname, "container-models.json");

// Root schemas; transitive imports inside transitional/ are loaded recursively.
const ROOT_XSDS = [
  "wml.xsd",
  "pml.xsd",
  "sml.xsd",
  "dml-main.xsd",
  "dml-chart.xsd",
  "dml-diagram.xsd",
  "dml-wordprocessingDrawing.xsd",
  "dml-spreadsheetDrawing.xsd",
  "dml-picture.xsd",
  "dml-lockedCanvas.xsd",
  "dml-chartDrawing.xsd",
  "vml-main.xsd",
  "vml-officeDrawing.xsd",
  "vml-wordprocessingDrawing.xsd",
  "vml-spreadsheetDrawing.xsd",
  "vml-presentationDrawing.xsd",
  "shared-math.xsd",
];

// Conventional output prefixes per target namespace. The XSD files themselves
// declare whatever prefixes they like; the ecosystem (and this repo) uses these,
// and local element declarations carry no prefix in the source.
const NS_PREFIX: Record<string, string> = {
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main": "w",
  "http://schemas.openxmlformats.org/presentationml/2006/main": "p",
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main": "",
  "http://schemas.openxmlformats.org/drawingml/2006/main": "a",
  "http://schemas.openxmlformats.org/drawingml/2006/chart": "c",
  "http://schemas.openxmlformats.org/drawingml/2006/diagram": "dgm",
  "http://schemas.openxmlformats.org/drawingml/2006/picture": "pic",
  "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas": "lc",
  "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing": "cdr",
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing": "wp",
  "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing": "xdr",
  "http://schemas.openxmlformats.org/officeDocument/2006/math": "m",
  "urn:schemas-microsoft-com:vml": "v",
  "urn:schemas-microsoft-com:office:office": "o",
  "urn:schemas-microsoft-com:office:word": "w10",
  "urn:schemas-microsoft-com:office:excel": "x",
  "urn:schemas-microsoft-com:office:powerpoint": "pvml",
  "http://schemas.openxmlformats.org/schemaLibrary/2006/main": "sl",
  "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes": "s",
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships": "r",
  "http://schemas.openxmlformats.org/markup-compatibility/2006": "mc",
};

interface Slot {
  elements?: string[];
  any?: string[];
  min: number;
  max: number | "unbounded";
  pick?: "one" | "any";
  // choice arms, present only when some arm expands to several elements
  // (a group ref): members combine freely within an arm, exclusively across
  arms?: string[][];
  // group-ref with unbounded repeat over a multi-slot body cannot be expressed
  // by inline slots without losing the "repeat the whole sequence" semantics
  approxRepeat?: true;
}

interface Container {
  ct: string[];
  slots: Slot[];
  // union of every CT variant's slot members (w:sdtContent block/run/cell/row
  // variants are context-dependent) — set-level gate input, ordered gates use slots
  variantElements?: string[];
}

// ── Minimal XML tokenizer ────────────────────────────────────────────────────
// The transitional XSDs are machine-generated, comment-light, CDATA-free XML;
// a small stack tokenizer is sufficient (no repo dependency).

interface XNode {
  tag: string; // local name, prefix stripped
  attrs: Record<string, string>;
  children: XNode[];
}

function parseXml(text: string): XNode {
  // strip comments and processing instructions
  const src = text.replace(/<!--[\s\S]*?-->/g, "").replace(/<\?[\s\S]*?\?>/g, "");
  const root: XNode = { tag: "#root", attrs: {}, children: [] };
  const stack: XNode[] = [root];
  let i = 0;
  const len = src.length;
  while (i < len) {
    const lt = src.indexOf("<", i);
    if (lt < 0) break;
    if (src[lt + 1] === "/") {
      const gt = src.indexOf(">", lt);
      stack.pop();
      i = gt + 1;
      continue;
    }
    // scan to tag end, honoring quoted attribute values (patterns may contain ">")
    let j = lt + 1;
    let quote: string | null = null;
    while (j < len) {
      const ch = src[j]!;
      if (quote) {
        if (ch === quote) quote = null;
      } else if (ch === '"' || ch === "'") {
        quote = ch;
      } else if (ch === ">") {
        break;
      }
      j++;
    }
    const raw = src.slice(lt + 1, j);
    const selfClose = raw.endsWith("/");
    const body = selfClose ? raw.slice(0, -1) : raw;
    const nameMatch = body.match(/^([^\s/]+)/);
    if (!nameMatch) {
      i = j + 1;
      continue;
    }
    const qname = nameMatch[1]!;
    const local = qname.includes(":") ? qname.slice(qname.indexOf(":") + 1) : qname;
    const attrs: Record<string, string> = {};
    const attrRe = /([^\s=]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
    let m: RegExpExecArray | null;
    while ((m = attrRe.exec(body)) !== null) {
      attrs[m[1]!] = m[3] ?? m[4] ?? "";
    }
    const node: XNode = { tag: local, attrs, children: [] };
    stack[stack.length - 1]!.children.push(node);
    if (!selfClose) stack.push(node);
    i = j + 1;
  }
  return root;
}

// ── Schema registry ──────────────────────────────────────────────────────────

interface SchemaFile {
  file: string;
  targetNs: string;
  prefix: string; // conventional output prefix for local declarations
  complexTypes: Map<string, XNode>; // CT name → complexType node
  groups: Map<string, XNode>; // group name (prefixed, as referenced) → group node
  // imports: referenced prefix → schema file basename
  importPrefixes: Map<string, string>;
}

const schemas = new Map<string, SchemaFile>(); // basename → SchemaFile
const warnings: string[] = [];

function loadSchema(file: string): SchemaFile {
  const cached = schemas.get(file);
  if (cached) return cached;

  const text = fs.readFileSync(path.join(SCHEMA_DIR, file), "utf-8");
  const root = parseXml(text);
  const schemaEl = root.children.find((c) => c.tag === "schema")!;

  const targetNs = schemaEl.attrs.targetNamespace ?? "";
  const prefix = NS_PREFIX[targetNs];
  if (prefix === undefined) {
    throw new Error(`${file}: no conventional prefix for targetNamespace ${targetNs}`);
  }

  // xmlns declarations → importable namespace prefixes
  const importPrefixes = new Map<string, string>();
  for (const [k, v] of Object.entries(schemaEl.attrs)) {
    if (!k.startsWith("xmlns:")) continue;
    const p = k.slice(6);
    // a prefix maps to an import when some xsd:import declares that namespace
    if (schemaEl.children.some((c) => c.tag === "import" && c.attrs.namespace === v)) {
      const imp = schemaEl.children.find((c) => c.tag === "import" && c.attrs.namespace === v);
      if (imp?.attrs.schemaLocation) importPrefixes.set(p, path.basename(imp.attrs.schemaLocation));
    }
  }

  const complexTypes = new Map<string, XNode>();
  const groups = new Map<string, XNode>();
  // groups are keyed by prefixed name as they appear in refs, so cross-file
  // refs resolve through the referencing file's prefix map
  for (const c of schemaEl.children) {
    if (c.tag === "complexType" && c.attrs.name) complexTypes.set(c.attrs.name, c);
    if (c.tag === "group" && c.attrs.name) {
      const qname = prefix ? `${prefix}:${c.attrs.name}` : c.attrs.name;
      groups.set(qname, c);
    }
  }

  const entry: SchemaFile = { file, targetNs, prefix, complexTypes, groups, importPrefixes };
  schemas.set(file, entry);

  // recurse into imports that live inside transitional/
  for (const loc of importPrefixes.values()) {
    if (schemas.has(loc)) continue;
    if (fs.existsSync(path.join(SCHEMA_DIR, loc))) loadSchema(loc);
    else warnings.push(`${file}: import ${loc} outside transitional/ — skipped`);
  }
  return entry;
}

// ── Particle flattening ──────────────────────────────────────────────────────

function cardinality(node: XNode): { min: number; max: number | "unbounded" } {
  const min = node.attrs.minOccurs !== undefined ? Number(node.attrs.minOccurs) : 1;
  const max =
    node.attrs.maxOccurs === "unbounded"
      ? ("unbounded" as const)
      : node.attrs.maxOccurs !== undefined
        ? Number(node.attrs.maxOccurs)
        : 1;
  return { min, max };
}

function mergeMax(a: number | "unbounded", b: number | "unbounded"): number | "unbounded" {
  if (a === "unbounded" || b === "unbounded") return "unbounded";
  return Math.max(a, b);
}

/**
 * Flatten a particle node (sequence/choice/all/group-ref/element/any) into
 * ordered slots. `refMin/refMax` carry the cardinality of the referencing
 * group-ref so choice groups expand with the right repeat semantics.
 */
function flatten(
  node: XNode,
  sf: SchemaFile,
  refMin: number,
  refMax: number | "unbounded",
  groupDepth = 0,
): Slot[] {
  if (groupDepth > 16) return []; // cycle guard

  if (node.tag === "element") {
    const { min, max } = cardinality(node);
    // same-file refs carry no prefix in the source (vml's ref="path") — they
    // still resolve into the owning schema's namespace
    const name = node.attrs.ref ? qualify(node.attrs.ref, sf) : qualify(node.attrs.name ?? "", sf);
    if (!name) return [];
    return [{ elements: [name], min, max }];
  }

  if (node.tag === "any") {
    const { min, max } = cardinality(node);
    const ns = (node.attrs.namespace ?? "##any").split(/\s+/);
    return [{ any: ns, min, max }];
  }

  if (node.tag === "group") {
    // group definition body — resolve through the owning schema; local member
    // declarations qualify with the *defining* file's prefix (a:EG_Media
    // members are a:audioCd even when referenced from pml)
    const resolved = resolveGroup(node, sf);
    if (!resolved) return [];
    const [target, owner] = resolved;
    const { min, max } = cardinality(node);
    return flattenChildren(target, owner, groupDepth + 1).map((s) =>
      // the ref's own cardinality wraps the group body: an optional ref makes
      // every inner slot optional, an unbounded ref makes pick groups "any"
      min === 1 && max === 1 ? s : scaleSlot(s, min, max, target),
    );
  }

  if (node.tag === "choice") {
    // each direct child particle is one ARM of the choice. A group-ref arm
    // expands to its members, which combine freely WITHIN the arm — `arms`
    // records that boundary so gates test exclusivity ACROSS arms only
    // (c:dLbls: delete XOR the shared-settings group, never two shared
    // settings). A nested choice flattens into its parent arm; exclusivity
    // inside it degrades to the arm union (rare, approximate).
    const { min, max } = cardinality(node);
    const pick = max === "unbounded" || refMax === "unbounded" ? "any" : "one";
    const elements: string[] = [];
    const arms: string[][] = [];
    const wildcards: Slot[] = [];
    let slotMin = min;
    // the slot may repeat as often as the choice itself (or its referencing
    // ref) allows — unbounded there is what makes pick "any" repeatable
    let slotMax: number | "unbounded" = mergeMax(max, refMax);
    for (const child of node.children) {
      if (!isParticle(child)) continue;
      const armElements: string[] = [];
      for (const s of flatten(child, sf, 1, 1, groupDepth + 1)) {
        if (s.elements) armElements.push(...s.elements);
        else wildcards.push(s);
        slotMin = Math.min(slotMin, s.min);
        slotMax = mergeMax(slotMax, s.max);
      }
      if (armElements.length > 0) {
        arms.push(armElements);
        elements.push(...armElements);
      }
    }
    // single-element arms carry no extra information — the flat member list
    // already expresses pairwise exclusivity; only record multi-element arms
    const hasGroupArm = arms.some((a) => a.length > 1);
    if (elements.length === 0) return wildcards;
    const slot: Slot = {
      elements,
      min: slotMin,
      max: slotMax,
      pick,
      ...(hasGroupArm ? { arms } : {}),
    };
    return [slot, ...wildcards];
  }

  if (node.tag === "sequence" || node.tag === "all") {
    const { min, max } = cardinality(node);
    const inner = flattenChildren(node, sf, groupDepth + 1);
    return min === 1 && max === 1 ? inner : inner.map((s) => scaleSlot(s, min, max, node));
  }
  return [];
}

function isParticle(node: XNode): boolean {
  return ["element", "group", "choice", "sequence", "all", "any"].includes(node.tag);
}

// Wrapping a group/choice body in the referencing particle's cardinality:
// optional ref → all inner slots optional; unbounded ref → pick groups may
// repeat ("any"). A bounded ref around a multi-particle body cannot
// repeat-the-sequence when inlined — flagged as approximate.
function scaleSlot(s: Slot, min: number, max: number | "unbounded", owner: XNode): Slot {
  const childCount = owner.children.filter(isParticle).length;
  const out: Slot = {
    ...s,
    min: min === 0 ? 0 : Math.min(s.min, min),
    max: mergeMax(s.max, max),
  };
  if (max === "unbounded" && s.pick) out.pick = "any";
  if (childCount > 1 && max === "unbounded") {
    out.approxRepeat = true;
    warnings.push(`approxRepeat: multi-particle body under unbounded ref (${describe(owner)})`);
  }
  return out;
}

function describe(node: XNode): string {
  if (node.attrs.name) return node.attrs.name;
  if (node.attrs.ref) return `ref ${node.attrs.ref}`;
  return node.tag;
}

function flattenChildren(container: XNode, sf: SchemaFile, groupDepth: number): Slot[] {
  const out: Slot[] = [];
  for (const child of container.children) {
    if (!isParticle(child)) continue;
    out.push(...flatten(child, sf, 1, 1, groupDepth));
  }
  return out;
}

function qualify(name: string, sf: SchemaFile): string {
  if (name.includes(":")) return name; // already qualified as written
  return sf.prefix ? `${sf.prefix}:${name}` : name;
}

// Resolve a group definition: local first, then through import prefixes.
// Resolve a group definition to [node, owning schema]: ref="p:EG_..." looks in
// the schema imported under prefix p; member declarations qualify with the
// *defining* file's prefix (a:EG_Media members are a:audioCd even from pml).
function resolveGroup(refNode: XNode, sf: SchemaFile): [XNode, SchemaFile] | null {
  const ref = refNode.attrs.ref ?? refNode.attrs.name;
  if (!ref) return null;
  if (ref.includes(":")) {
    const p = ref.slice(0, ref.indexOf(":"));
    const targetFile = sf.importPrefixes.get(p);
    if (targetFile && schemas.has(targetFile)) {
      const owner = schemas.get(targetFile)!;
      const g = owner.groups.get(ref);
      if (g) return [g, owner];
    }
    // fall back: any schema whose prefixed group matches
    for (const s of schemas.values()) {
      const g = s.groups.get(ref);
      if (g) return [g, s];
    }
    warnings.push(`${sf.file}: unresolved group ref ${ref}`);
    return null;
  }
  const local = sf.groups.get(qualify(ref, sf));
  return local ? [local, sf] : null;
}

// Resolve a complexType's content model. complexContent/extension chains the
// base type's slots before the extension body (XSD extension semantics);
// restriction and simpleContent carry their own (or no) particle tree.
function complexTypeSlots(ctName: string, sf: SchemaFile, seen: Set<string>): Slot[] {
  if (seen.has(ctName)) return [];
  seen.add(ctName);

  let node: XNode | undefined = sf.complexTypes.get(ctName);
  let owner: SchemaFile = sf;
  if (!node) {
    for (const s of schemas.values()) {
      const found = s.complexTypes.get(ctName);
      if (found) {
        node = found;
        owner = s;
        break;
      }
    }
  }
  if (!node) return [];

  const complexContent = node.children.find((c) => c.tag === "complexContent");
  if (complexContent) {
    const body = complexContent.children.find(
      (c) => c.tag === "extension" || c.tag === "restriction",
    );
    if (!body) return [];
    if (body.tag === "extension") {
      const base = body.attrs.base ?? "";
      const baseSlots = base && !base.includes(":") ? complexTypeSlots(base, owner, seen) : [];
      return [...baseSlots, ...flattenChildren(body, owner, 0)];
    }
    return flattenChildren(body, owner, 0);
  }
  if (node.children.some((c) => c.tag === "simpleContent")) return [];
  return flattenChildren(node, owner, 0);
}

// ── Element table assembly ───────────────────────────────────────────────────

const containers = new Map<string, Container>();
const namespaces: Record<string, string> = {};

// Richness for "most complete model wins": slot count first, total member
// count as tie-breaker (w:ins has a 1-slot CT_MathCtrlIns variant vs the
// 1-slot pick-any CT_RunTrackChange carrying every run-level element).
const richness = (slots: Slot[]) =>
  slots.length * 1000 + slots.reduce((n, s) => n + (s.elements?.length ?? 0), 0);

function addElement(elementQName: string, ctName: string, slots: Slot[]) {
  const unionInto = (target: Container) => {
    for (const s of slots) {
      for (const e of s.elements ?? []) {
        if (!target.variantElements) target.variantElements = [];
        if (!target.variantElements.includes(e)) target.variantElements.push(e);
      }
    }
  };
  const existing = containers.get(elementQName);
  if (existing) {
    if (!existing.ct.includes(ctName)) existing.ct.push(ctName);
    // several elements share a base and a full complexType (CT_SectPrBase vs
    // CT_SectPr); keep the most complete slot list for gate purposes
    if (richness(slots) > richness(existing.slots)) existing.slots = slots;
    unionInto(existing);
    return;
  }
  const fresh: Container = { ct: [ctName], slots };
  containers.set(elementQName, fresh);
  unionInto(fresh);
}

// A type reference may cross namespaces (pml's `<xsd:element name="to"
// type="a:CT_Color"/>` inside CT_TLAnimateColorBehavior): resolve the prefix
// through the referencing file's imports to the owning schema so those
// element models land in the table instead of being dropped.
function resolveTypeRef(type: string, sf: SchemaFile): [SchemaFile, string] | undefined {
  if (!type.includes(":")) return [sf, type];
  const px = type.slice(0, type.indexOf(":"));
  const local = type.slice(type.indexOf(":") + 1);
  const file = sf.importPrefixes.get(px);
  const target = file ? schemas.get(file) : undefined;
  if (target?.complexTypes.has(local)) return [target, local];
  return undefined;
}

function collectFromSchema(sf: SchemaFile) {
  const schemaEl = parseXml(fs.readFileSync(path.join(SCHEMA_DIR, sf.file), "utf-8")).children.find(
    (c) => c.tag === "schema",
  )!;

  const walk = (node: XNode) => {
    for (const child of node.children) {
      // element declarations with a named complex type produce containers
      if (child.tag === "element") {
        const qname = child.attrs.ref ?? qualify(child.attrs.name ?? "", sf);
        const type = child.attrs.type;
        if (qname && type) {
          const resolved = resolveTypeRef(type, sf);
          if (resolved) {
            const [typeSf, typeName] = resolved;
            const slots = complexTypeSlots(typeName, typeSf, new Set());
            if (slots.length > 0 || typeSf.complexTypes.has(typeName)) {
              addElement(qname, typeName, slots);
            }
          }
        }
      }
      walk(child);
    }
  };
  walk(schemaEl);
}

// ── Main ─────────────────────────────────────────────────────────────────────

function build(): string {
  for (const root of ROOT_XSDS) loadSchema(root);
  for (const sf of schemas.values()) collectFromSchema(sf);

  for (const [uri, prefix] of Object.entries(NS_PREFIX)) {
    if ([...schemas.values()].some((s) => s.targetNs === uri)) {
      namespaces[prefix] = uri;
    }
  }

  const orderedContainers: Record<string, Container> = {};
  for (const key of [...containers.keys()].sort()) {
    const c = containers.get(key)!;
    if (c.variantElements) c.variantElements.sort();
    orderedContainers[key] = c;
  }

  const payload = {
    generatedFrom: "ooxml-schemas/transitional",
    roots: ROOT_XSDS,
    namespaces,
    containers: orderedContainers,
  };
  return JSON.stringify(payload, null, 2) + "\n";
}

/**
 * The repo formatter (oxfmt) rewraps the generated JSON (folds short arrays
 * onto one line), so the drift gate compares parsed values, not text —
 * `vp check --fix` and `models:generate` must not fight over formatting.
 */
function sameJson(a: string, b: string): boolean {
  if (a === b) return true;
  try {
    return JSON.stringify(JSON.parse(a)) === JSON.stringify(JSON.parse(b));
  } catch {
    return false;
  }
}

function main() {
  const json = build();
  const check = process.argv.includes("--check");

  if (check) {
    const committed = fs.existsSync(OUT_FILE) ? fs.readFileSync(OUT_FILE, "utf-8") : "";
    if (sameJson(committed, json)) {
      console.log(`container-models.json up to date (${containers.size} containers)`);
      return;
    }
    console.error("container-models.json is stale — regenerate with:");
    console.error("  pnpm tsx scripts/xsd-content-model.ts");
    process.exitCode = 1;
    return;
  }

  fs.writeFileSync(OUT_FILE, json);
  const slotCount = [...containers.values()].reduce((n, c) => n + c.slots.length, 0);
  console.log(
    `container-models.json: ${containers.size} containers, ${slotCount} slots, ${schemas.size} schemas loaded`,
  );
  if (warnings.length > 0) {
    console.log(`warnings (${warnings.length}):`);
    for (const w of [...new Set(warnings)].slice(0, 20)) console.log(`  - ${w}`);
  }

  // sanity spot-checks against hand-verified XSD orderings; w:rPr aggregates
  // CT_RPr/CT_ParaRPr/…Original variants, so assert on invariant members
  const rpr = containers.get("w:rPr");
  const anySlot = rpr?.slots.find((s) => s.pick === "any");
  const okRpr =
    rpr !== undefined &&
    anySlot !== undefined &&
    anySlot.elements?.[0] === "w:rStyle" &&
    anySlot.elements.includes("w:sz") &&
    rpr.slots.at(-1)?.elements?.[0] === "w:rPrChange";
  const sectPr = containers.get("w:sectPr");
  const okSectPr =
    sectPr !== undefined &&
    sectPr.slots.length >= 3 &&
    sectPr.slots[0]?.pick === "one" &&
    sectPr.slots[0].elements?.[0] === "w:headerReference" &&
    sectPr.slots.at(-1)?.elements?.[0] === "w:sectPrChange";
  console.log(`spot-checks: w:rPr ${okRpr ? "ok" : "FAIL"}, w:sectPr ${okSectPr ? "ok" : "FAIL"}`);
  if (!okRpr || !okSectPr) process.exitCode = 1;
}

main();
