/**
 * Skeleton schemas for the generate tools' `inputSchema`.
 *
 * The full format schema (docx: ~675 KB) is far beyond what any provider
 * accepts in a tool definition. The skeleton is derived at runtime from the
 * real schema so it can never drift from it: top-level fields stay verbatim,
 * the section/slide/worksheet chain expands one level, tagged-union wrapper
 * keys (the `{ paragraph: {...} }` / `{ shape: {...} }` convention) expand
 * into a scannable key list, and everything deeper becomes a stub that names
 * the definition to look up via the schema-lookup tool. The output is fully
 * inlined — no `$ref`/`definitions` — for maximum provider compatibility.
 *
 * A skeleton deliberately accepts more than the full schema (stubs allow any
 * value); the authoritative ajv validation runs inside the tool's `execute`.
 *
 * @module
 */

import { SCHEMAS, type DocumentType, type JsonSchema } from "./schemas";

/** Definitions expanded in the skeleton, outermost first. */
const SKELETON_SPINE: Record<DocumentType, readonly string[]> = {
  docx: ["DocumentOptions", "SectionOptions"],
  pptx: ["PresentationOptions", "SlideOptions"],
  xlsx: ["WorkbookOptions", "WorksheetOptions", "RowOptions", "CellOptions"],
};

type Node = Record<string, unknown>;

function definitionsOf(type: DocumentType): Record<string, Node> {
  return SCHEMAS[type].definitions as Record<string, Node>;
}

/** A tagged-union wrapper: every branch is `{ title, properties: {<key>: …}, required: [<key>] }`. */
function isWrapperUnion(def: Node | undefined): boolean {
  const branches = def?.anyOf;
  if (!Array.isArray(branches) || branches.length === 0) return false;
  return branches.every((branch) => {
    const b = branch as Node;
    const required = b.required as string[] | undefined;
    const properties = b.properties as Record<string, unknown> | undefined;
    return (
      typeof b.title === "string" &&
      Array.isArray(required) &&
      required.length === 1 &&
      properties !== undefined &&
      Object.keys(properties).length === 1
    );
  });
}

function buildStub(name: string): Node {
  return {
    description: `"${name}" stub — fetch its fields with the office-open-schema-lookup tool.`,
    $comment: `office-open-stub:${name}`,
  };
}

/** Collapse an inline (non-$ref) object schema that is too large to be self-explanatory. */
function collapseInline(node: Node, format: DocumentType, context: string): Node {
  const description = typeof node.description === "string" ? `${node.description}\n\n` : "";
  return {
    type: "object",
    description:
      `${description}Collapsed inline object (${context}) — call the office-open-schema-lookup tool ` +
      `with { type: "${format}" } for details.`,
  };
}

function convertProperty(
  prop: Node,
  format: DocumentType,
  spine: readonly string[],
  visited: ReadonlySet<string>,
): Node {
  const ref = typeof prop.$ref === "string" ? prop.$ref : undefined;
  if (ref) {
    const name = decodeURIComponent(ref.slice("#/definitions/".length));
    const target = definitionsOf(format)[name];
    if (visited.has(name)) return buildStub(name);
    if (spine.includes(name)) {
      return convertDefinition(name, target, format, new Set([...visited, name]));
    }
    if (isWrapperUnion(target)) {
      // visited must thread through wrapper expansion: unions can reference
      // themselves (SectionChild.sdt.children → SectionChild).
      return convertWrapperUnion(target, format, new Set([...visited, name]));
    }
    return buildStub(name);
  }

  if (Array.isArray(prop.anyOf) || Array.isArray(prop.oneOf)) {
    const key = Array.isArray(prop.anyOf) ? "anyOf" : "oneOf";
    const branches = (prop[key] as Node[]).map((branch) =>
      convertProperty(branch, format, spine, visited),
    );
    const out: Node = { [key]: branches };
    if (typeof prop.description === "string") out.description = prop.description;
    // A wrapper key list (titled branches) is core skeleton information;
    // any other oversized union collapses to a lookup pointer.
    const isWrapperList = branches.some((branch) => branch.title !== undefined);
    if (!isWrapperList && JSON.stringify(out).length >= 400) {
      return collapseInline(prop, format, "inline union");
    }
    return out;
  }

  if (prop.type === "array" && prop.items && typeof prop.items === "object") {
    return {
      type: "array",
      ...(typeof prop.description === "string" ? { description: prop.description } : {}),
      items: convertProperty(prop.items as Node, format, spine, visited),
    };
  }

  // Scalars, enums, and small inline objects pass through untouched — but a
  // $ref must never escape into this $ref-less skeleton, so inline schemas
  // that reference other definitions are recursed through their containers
  // first and then run through the same size gate.
  const clone = structuredClone(prop) as Node;
  const text = JSON.stringify(clone);
  if (text.includes("#/definitions/")) {
    if (clone.properties && typeof clone.properties === "object") {
      clone.properties = Object.fromEntries(
        Object.entries(clone.properties as Record<string, Node>).map(([key, value]) => [
          key,
          convertProperty(value, format, spine, visited),
        ]),
      );
    }
    if (clone.items && typeof clone.items === "object") {
      clone.items = convertProperty(clone.items as Node, format, spine, visited);
    }
  }
  return JSON.stringify(clone).length < 400
    ? clone
    : collapseInline(prop, format, "large inline object");
}

function convertWrapperUnion(def: Node, format: DocumentType, visited: ReadonlySet<string>): Node {
  const out: Node = {};
  if (typeof def.description === "string") out.description = def.description;
  out.anyOf = (def.anyOf as Node[]).map((branch) => {
    const b = branch as Node;
    const properties = b.properties as Record<string, Node>;
    const key = Object.keys(properties)[0]!;
    const skeletonBranch: Node = {
      title: b.title,
      type: "object",
      properties: { [key]: convertProperty(properties[key]!, format, [], visited) },
      required: b.required,
    };
    if (typeof b.description === "string") skeletonBranch.description = b.description;
    return skeletonBranch;
  });
  return out;
}

function convertDefinition(
  name: string,
  def: Node,
  format: DocumentType,
  visited: ReadonlySet<string>,
): Node {
  const spine = SKELETON_SPINE[format].slice(SKELETON_SPINE[format].indexOf(name) + 1);
  const out: Node = {};
  if (typeof def.description === "string") out.description = def.description;
  if (typeof def.type === "string") out.type = def.type;
  const properties = def.properties as Record<string, Node> | undefined;
  if (properties) {
    const converted: Record<string, Node> = {};
    for (const [key, value] of Object.entries(properties)) {
      converted[key] = convertProperty(value, format, spine, visited);
    }
    out.properties = converted;
  }
  if (Array.isArray(def.required)) out.required = structuredClone(def.required);
  return out;
}

const skeletonCache = new Map<DocumentType, JsonSchema>();

/** Derive (and memoize) the skeleton input schema for a format. */
export function getSkeletonSchema(type: DocumentType): JsonSchema {
  let skeleton = skeletonCache.get(type);
  if (!skeleton) {
    const [root] = SKELETON_SPINE[type];
    if (!root) throw new Error(`no skeleton spine for ${type}`);
    const visited = new Set<string>([root]);
    skeleton = convertDefinition(root, definitionsOf(type)[root]!, type, visited) as JsonSchema;
    skeletonCache.set(type, skeleton);
  }
  return skeleton;
}
