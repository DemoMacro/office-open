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
  // group-ref with unbounded repeat over a multi-slot body cannot be expressed
  // by inline slots without losing the "repeat the whole sequence" semantics
  approxRepeat?: true;
}

interface Container {
  ct: string[];
  slots: Slot[];
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
    const name = node.attrs.ref ?? qualify(node.attrs.name ?? "", sf);
    if (!name) return [];
    return [{ elements: [name], min, max }];
  }

  if (node.tag === "any") {
    const { min, max } = cardinality(node);
    const ns = (node.attrs.namespace ?? "##any").split(/\s+/);
    return [{ any: ns, min, max }];
  }

  if (node.tag === "group") {
    // group definition body — resolve through the owning schema
    const target = resolveGroup(node, sf);
    if (!target) return [];
    const { min, max } = cardinality(node);
    return flattenChildren(target, sf, groupDepth + 1).map((s) =>
      // the ref's own cardinality wraps the group body: an optional ref makes
      // every inner slot optional, an unbounded ref makes pick groups "any"
      min === 1 && max === 1 ? s : scaleSlot(s, min, max, target),
    );
  }

  if (node.tag === "sequence" || node.tag === "choice" || node.tag === "all") {
    const { min, max } = cardinality(node);
    const inner = flattenChildren(node, sf, groupDepth + 1);
    const scaled = min === 1 && max === 1 ? inner : inner.map((s) => scaleSlot(s, min, max, node));
    if (node.tag === "choice" && scaled.length > 0) {
      // merge members into one pick slot; keep any-wildcards as their own slots
      const elements: string[] = [];
      const wildcards: Slot[] = [];
      for (const s of scaled) {
        if (s.elements) elements.push(...s.elements);
        else wildcards.push(s);
      }
      const slot: Slot =
        elements.length > 0
          ? {
              elements,
              min: scaled[0]!.min,
              max: scaled[0]!.max,
              pick: scaled[0]!.max === "unbounded" || refMax === "unbounded" ? "any" : "one",
            }
          : wildcards[0]!;
      return [slot, ...wildcards.slice(elements.length > 0 ? 0 : 1)];
    }
    return scaled;
  }
  return [];
}

// Wrapping a group/choice body in the referencing particle's cardinality:
// optional ref → all inner slots optional; unbounded ref → pick groups may
// repeat ("any"). A bounded ref around a multi-particle body cannot
// repeat-the-sequence when inlined — flagged as approximate.
function scaleSlot(s: Slot, min: number, max: number | "unbounded", owner: XNode): Slot {
  const childCount = owner.children.filter((c) =>
    ["element", "group", "choice", "sequence", "all", "any"].includes(c.tag),
  ).length;
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
    if (!["element", "group", "choice", "sequence", "all", "any"].includes(child.tag)) continue;
    out.push(...flatten(child, sf, 1, 1, groupDepth));
  }
  return out;
}

function qualify(name: string, sf: SchemaFile): string {
  if (name.includes(":")) return name; // already qualified as written
  return sf.prefix ? `${sf.prefix}:${name}` : name;
}

// Resolve a group definition: local first, then through import prefixes.
function resolveGroup(refNode: XNode, sf: SchemaFile): XNode | null {
  const ref = refNode.attrs.ref;
  if (!ref)
    return refNode.attrs.name ? (sf.groups.get(qualify(refNode.attrs.name, sf)) ?? null) : null;
  // ref="p:EG_..." → look in the schema imported under prefix p
  if (ref.includes(":")) {
    const p = ref.slice(0, ref.indexOf(":"));
    const targetFile = sf.importPrefixes.get(p);
    if (targetFile && schemas.has(targetFile)) {
      const g = schemas.get(targetFile)!.groups.get(ref);
      if (g) return g;
    }
    // fall back: any schema whose prefixed group matches
    for (const s of schemas.values()) {
      const g = s.groups.get(ref);
      if (g) return g;
    }
    warnings.push(`${sf.file}: unresolved group ref ${ref}`);
    return null;
  }
  return sf.groups.get(qualify(ref, sf)) ?? null;
}

// Same-schema-type resolution for complexType refs used by element type=...
function complexTypeSlots(ctName: string, sf: SchemaFile, seen: Set<string>): Slot[] {
  if (seen.has(ctName)) return [];
  seen.add(ctName);
  const local = sf.complexTypes.get(ctName);
  if (local) return flattenChildren(local, sf, 0);

  // cross-file: simple type names (ST_*) and types from imported schemas
  for (const s of schemas.values()) {
    const node = s.complexTypes.get(ctName);
    if (node) return flattenChildren(node, s, 0);
  }
  return [];
}

// ── Element table assembly ───────────────────────────────────────────────────

const containers = new Map<string, Container>();
const namespaces: Record<string, string> = {};

function addElement(elementQName: string, ctName: string, slots: Slot[]) {
  const existing = containers.get(elementQName);
  if (existing) {
    if (!existing.ct.includes(ctName)) existing.ct.push(ctName);
    // several elements share a base and a full complexType (CT_SectPrBase vs
    // CT_SectPr); keep the most complete slot list for gate purposes
    if (slots.length > existing.slots.length) existing.slots = slots;
    return;
  }
  containers.set(elementQName, { ct: [ctName], slots });
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
        if (qname && type && !type.includes(":")) {
          const slots = complexTypeSlots(type, sf, new Set());
          if (slots.length > 0 || sf.complexTypes.has(type)) addElement(qname, type, slots);
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
    orderedContainers[key] = containers.get(key)!;
  }

  const payload = {
    generatedFrom: "ooxml-schemas/transitional",
    roots: ROOT_XSDS,
    namespaces,
    containers: orderedContainers,
  };
  return JSON.stringify(payload, null, 2) + "\n";
}

function main() {
  const json = build();
  const check = process.argv.includes("--check");

  if (check) {
    const committed = fs.existsSync(OUT_FILE) ? fs.readFileSync(OUT_FILE, "utf-8") : "";
    if (committed === json) {
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
