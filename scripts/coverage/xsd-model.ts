/**
 * Structured XSD model — replaces the regex-based xsd.ts scanner.
 *
 * Parses each transitional schema with the real XML parser and resolves
 * global element declarations into typed models: attributes (direct,
 * attributeGroup, and base-type chain) plus direct child element refs.
 * This is the foundation for the per-QName ownership registry and the
 * runtime fixture matrix (coverage plan phase 1).
 *
 * Usage: pnpm exec tsx scripts/coverage/xsd-model.ts [--json <out>] [--summary]
 *
 * @module
 */

import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

import { parse, type Element } from "../../packages/xml/dist/index.mjs";

const ROOT_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const SCHEMA_DIR = path.join(ROOT_DIR, "ooxml-schemas", "transitional");

// ── Model ──

export interface XsdAttributeModel {
  name: string;
  type?: string;
  use?: string;
  default?: string;
}

export interface XsdChildModel {
  name: string;
  minOccurs?: string;
  maxOccurs?: string;
}

export interface XsdElementModel {
  /** Schema prefix (w, s, p, a, c, …) + local name. */
  qname: string;
  attributes: XsdAttributeModel[];
  children: XsdChildModel[];
}

export interface XsdSchemaModel {
  file: string;
  prefix: string;
  targetNamespace?: string;
  elements: Map<string, XsdElementModel>;
  namedTypes: Map<string, Element>;
  attributeGroups: Map<string, Element>;
  deprecated: Set<string>;
}

const FILE_PREFIX: Record<string, string> = {
  "wml.xsd": "w",
  "sml.xsd": "s",
  "pml.xsd": "p",
  "dml-main.xsd": "a",
  "dml-chart.xsd": "c",
  "dml-chartDrawing.xsd": "cd",
  "dml-diagram.xsd": "dgm",
  "dml-lockedCanvas.xsd": "lc",
  "dml-picture.xsd": "pic",
  "dml-spreadsheetDrawing.xsd": "xdr",
  "dml-wordprocessingDrawing.xsd": "wp",
  "shared-math.xsd": "m",
  "shared-relationshipReference.xsd": "r",
  "shared-commonSimpleTypes.xsd": "",
};

const SHARED_PREFIX: Record<string, string> = {
  "shared-commonSimpleTypes.xsd": "",
  "shared-math.xsd": "m",
  "shared-relationshipReference.xsd": "r",
};

// Deprecated element sets mirror scripts/coverage/xsd.ts so counts line up.
const DEPRECATED_ELEMENTS = new Set(["t", "u", "v", "w", "x", "y", "z"]);

// ── Helpers ──

function attrOf(el: Element | undefined, name: string): string | undefined {
  return el?.attributes?.[name];
}

function directChildren(el: Element, name: string): Element[] {
  return (el.elements ?? []).filter((child) => child.name === name);
}

function firstChild(el: Element, name: string): Element | undefined {
  return (el.elements ?? []).find((child) => child.name === name);
}

/** Attributes declared directly on a complexType/extension/attributeGroup. */
function collectAttributes(container: Element | undefined): XsdAttributeModel[] {
  if (!container) return [];
  const out: XsdAttributeModel[] = [];
  const walk = (node: Element): void => {
    for (const child of node.elements ?? []) {
      if (child.name === "xsd:attribute" || child.name === "xs:attribute") {
        const name = attrOf(child, "name") ?? attrOf(child, "ref");
        if (name) {
          out.push({
            name,
            type: attrOf(child, "type"),
            use: attrOf(child, "use"),
            default: attrOf(child, "default"),
          });
        }
      } else if (child.name === "xsd:attributeGroup" || child.name === "xs:attributeGroup") {
        const ref = attrOf(child, "ref");
        if (ref && !ref.includes(":")) {
          // Named attributeGroup resolution happens against the schema model.
          namedAttributeGroupRefs.push({ container: out, ref });
        } else if (ref) {
          const [prefix, local] = ref.split(":");
          const group = foreignAttributeGroups.get(`${prefix}:${local}`);
          if (group) out.push(...collectAttributes(group));
        }
      } else if (
        child.name === "xsd:complexContent" ||
        child.name === "xs:complexContent" ||
        child.name === "xsd:simpleContent" ||
        child.name === "xs:simpleContent"
      ) {
        walk(child);
      } else if (child.name === "xsd:extension" || child.name === "xs:extension") {
        // Base type attributes come first (OOXML chains one level deep).
        const base = attrOf(child, "base");
        if (base && !base.startsWith("xsd:") && !base.startsWith("xs:")) {
          const local = base.includes(":") ? base.split(":")[1] : base;
          const baseType = currentModel?.namedTypes.get(local);
          if (baseType) {
            const def =
              firstChild(baseType, "xsd:complexType") ?? firstChild(baseType, "xs:complexType");
            out.push(...collectAttributes(def));
          }
        }
        walk(child);
      } else {
        walk(child);
      }
    }
  };
  walk(container);
  return out;
}

/** Direct element children from sequence/choice/all containers. */
function collectChildren(container: Element | undefined): XsdChildModel[] {
  if (!container) return [];
  const out: XsdChildModel[] = [];
  const walk = (node: Element): void => {
    for (const child of node.elements ?? []) {
      if (child.name === "xsd:element" || child.name === "xs:element") {
        const ref = attrOf(child, "ref");
        const name = ref ?? attrOf(child, "name");
        if (name) {
          out.push({
            name: name.includes(":") ? name.split(":")[1] : name,
            minOccurs: attrOf(child, "minOccurs"),
            maxOccurs: attrOf(child, "maxOccurs"),
          });
        }
      } else if (child.name !== "xsd:attribute" && child.name !== "xs:attribute") {
        walk(child);
      }
    }
  };
  const root =
    firstChild(container, "xsd:complexType") ??
    firstChild(container, "xs:complexType") ??
    firstChild(container, "xsd:complexContent") ??
    firstChild(container, "xs:complexContent") ??
    container;
  walk(root);
  return out;
}

// Module-level resolution state (single-threaded CLI).
let currentModel: XsdSchemaModel | undefined;
const namedAttributeGroupRefs: { container: XsdAttributeModel[]; ref: string }[] = [];
const foreignAttributeGroups = new Map<string, Element>();

// ── Parsing ──

export function parseXsdStructured(file: string): XsdSchemaModel {
  const content = fs.readFileSync(path.join(SCHEMA_DIR, file), "utf-8");
  const doc = parse(content).elements?.[0];
  if (!doc) throw new Error(`${file}: no root element`);
  const prefix = FILE_PREFIX[file] ?? SHARED_PREFIX[file] ?? "";
  const model: XsdSchemaModel = {
    file,
    prefix,
    targetNamespace: attrOf(doc, "targetNamespace"),
    elements: new Map(),
    namedTypes: new Map(),
    attributeGroups: new Map(),
    deprecated: DEPRECATED_ELEMENTS,
  };
  currentModel = model;

  for (const child of doc.elements ?? []) {
    if (child.name === "xsd:complexType" || child.name === "xs:complexType") {
      const name = attrOf(child, "name");
      if (name) model.namedTypes.set(name, child);
    } else if (child.name === "xsd:attributeGroup" || child.name === "xs:attributeGroup") {
      const name = attrOf(child, "name");
      if (name) model.attributeGroups.set(name, child);
    } else if (child.name === "xsd:import" || child.name === "xs:import") {
      const ns = attrOf(child, "namespace");
      const location = attrOf(child, "schemaLocation");
      if (ns && location) {
        const importedFile = path.basename(location);
        const importedPrefix =
          FILE_PREFIX[importedFile] ??
          SHARED_PREFIX[importedFile] ??
          // Transitional uses per-import prefixes (r, m, dcterms, …).
          importedFile.replace("shared-", "").replace(".xsd", "").slice(0, 2);
        try {
          const imported = parseXsdShallow(importedFile);
          for (const [name, group] of imported.attributeGroups)
            foreignAttributeGroups.set(`${importedPrefix}:${name}`, group);
        } catch {
          // Missing import files are tolerated — attributes from them are skipped.
        }
      }
    }
  }

  // Collect every element declaration (global and local) — documents carry
  // both, and coverage must model any element that can appear in a part.
  const collectElementDeclarations = (node: Element): void => {
    for (const child of node.elements ?? []) {
      if (child.name === "xsd:element" || child.name === "xs:element") {
        const name = attrOf(child, "name");
        if (name && !DEPRECATED_ELEMENTS.has(name) && !model.elements.has(name)) {
          const qname = prefix ? `${prefix}:${name}` : name;
          const typeAttr = attrOf(child, "type");
          const inlineType =
            firstChild(child, "xsd:complexType") ?? firstChild(child, "xs:complexType");
          let attributes: XsdAttributeModel[] = [];
          let children: XsdChildModel[] = [];
          if (inlineType) {
            attributes = collectAttributes(inlineType);
            children = collectChildren(inlineType);
          } else if (typeAttr && !typeAttr.startsWith("xsd:") && !typeAttr.startsWith("xs:")) {
            const local = typeAttr.includes(":") ? typeAttr.split(":")[1] : typeAttr;
            const named = model.namedTypes.get(local);
            if (named) {
              attributes = collectAttributes(named);
              children = collectChildren(named);
            }
          }
          model.elements.set(name, { qname, attributes, children });
        }
      }
      collectElementDeclarations(child);
    }
  };
  collectElementDeclarations(doc);

  // Second pass: attributeGroup refs resolved now that all groups are known.
  for (const { container, ref } of namedAttributeGroupRefs) {
    const local = ref.includes(":") ? ref.split(":")[1] : ref;
    const group = model.attributeGroups.get(local);
    if (group) container.push(...collectAttributes(group));
  }

  return model;
}

/** Parse only the type/group declarations (for import resolution). */
function parseXsdShallow(file: string): { attributeGroups: Map<string, Element> } {
  const content = fs.readFileSync(path.join(SCHEMA_DIR, file), "utf-8");
  const doc = parse(content).elements?.[0];
  if (!doc) return { attributeGroups };
  const attributeGroups = new Map<string, Element>();
  for (const child of doc.elements ?? []) {
    if (child.name === "xsd:attributeGroup" || child.name === "xs:attributeGroup") {
      const name = attrOf(child, "name");
      if (name) attributeGroups.set(name, child);
    }
  }
  return { attributeGroups };
}

// ── CLI ──

interface CliOptions {
  json?: string;
  summary: boolean;
  positional: string[];
}

function parseArgs(argv: string[]): CliOptions {
  const options: CliOptions = { summary: false, positional: [] };
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === "--json") options.json = argv[++i];
    else if (argv[i] === "--summary") options.summary = true;
    else options.positional.push(argv[i]);
  }
  return options;
}

function main(): void {
  const options = parseArgs(process.argv.slice(2));
  const files = options.positional.length
    ? options.positional
    : fs
        .readdirSync(SCHEMA_DIR)
        .filter((f) => f.endsWith(".xsd"))
        .sort();

  const schemas = files.map((f) => parseXsdStructured(f));
  let totalElements = 0;
  let totalAttributes = 0;
  const report: Record<string, { elements: number; attributes: number }> = {};

  for (const schema of schemas) {
    let attributeCount = 0;
    for (const element of schema.elements.values()) attributeCount += element.attributes.length;
    totalElements += schema.elements.size;
    totalAttributes += attributeCount;
    report[schema.file] = { elements: schema.elements.size, attributes: attributeCount };
  }

  if (options.json) {
    const out: Record<string, unknown> = {};
    for (const schema of schemas) {
      out[schema.file] = {
        prefix: schema.prefix,
        elements: Object.fromEntries(schema.elements),
      };
    }
    fs.mkdirSync(path.dirname(options.json), { recursive: true });
    fs.writeFileSync(options.json, `${JSON.stringify(out, null, 2)}\n`);
    console.log(`xsd model written to ${options.json}`);
  }

  if (options.summary || !options.json) {
    for (const [file, counts] of Object.entries(report))
      console.log(`  ${file}: ${counts.elements} elements, ${counts.attributes} attributes`);
    console.log(`structured total: ${totalElements} elements, ${totalAttributes} attributes`);
  }
}

const invokedDirectly = process.argv[1]?.endsWith("xsd-model.ts");
if (invokedDirectly) main();
