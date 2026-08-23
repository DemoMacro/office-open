/**
 * Type-definition rendering for schema slices.
 *
 * The AI-facing layer renders slices as type definitions instead of JSON
 * Schema: type definitions are a lossless compression of the same information
 * (BoundaryML's "type-definition prompting"), the tokenizers of current
 * models are already tuned to type syntax, and `?`-suffix optionality keeps
 * each field next to its constraint instead of a `required` list many tokens
 * away. Measured on real slices: 38-44% fewer tokens than the JSON form.
 *
 * Validation never goes through this format — the authoritative ajv gate
 * compiles the full draft-07 schemas.
 *
 * @module
 */

import type { DocumentType, JsonSchema } from "./schemas";

type Node = Record<string, unknown>;

/** Scalar-shorthand names for JSON Schema type keywords. */
const SCALAR_NAMES: Record<string, string> = {
  string: "string",
  number: "number",
  integer: "number",
  boolean: "boolean",
  null: "null",
};

function firstLine(text: string): string {
  return text.split("\n")[0]!.trim();
}

/** Compact a description to its first line for trailing comments. */
function trailingComment(node: Node): string {
  return typeof node.description === "string" && node.description.length > 0
    ? ` // ${firstLine(node.description)}`
    : "";
}

/**
 * Comment block for a definition header: the full multi-line description —
 * later lines often carry unit and encoding contracts that must survive.
 */
function headerComment(def: Node): string {
  if (typeof def.description !== "string" || def.description.length === 0) return "";
  const lines = def.description.split("\n").filter((l) => l.trim().length > 0);
  return ` // ${lines.join(" · ")}`;
}

/** Render a property/branch schema node as a compact type expression. */
function renderType(node: Node | undefined): string {
  if (!node || typeof node !== "object") return "unknown";
  if (typeof node.$ref === "string") {
    return decodeURIComponent(node.$ref.slice("#/definitions/".length));
  }
  const branches = (node.anyOf ?? node.oneOf) as Node[] | undefined;
  if (Array.isArray(branches)) return branches.map((b) => renderType(b)).join(" | ");
  if (Array.isArray(node.enum)) {
    return node.enum.map((v) => JSON.stringify(v)).join(" | ");
  }
  if (node.type === "array") {
    const inner = renderType(node.items as Node | undefined);
    return /^"[^"]*"|[A-Za-z0-9_.]+(\[\])?$/.test(inner) ? `${inner}[]` : `(${inner})[]`;
  }
  if (typeof node.type === "string" && node.type in SCALAR_NAMES) {
    return SCALAR_NAMES[node.type]!;
  }
  if (node.type === "object") {
    // inline object: single line, sub-descriptions dropped (the def they
    // usually point at carries them; keeping every line short is the point)
    const required = new Set(Array.isArray(node.required) ? (node.required as string[]) : []);
    const props = Object.entries((node.properties ?? {}) as Record<string, Node>).map(
      ([key, value]) => `${key}${required.has(key) ? "" : "?"}: ${renderType(value)}`,
    );
    return props.length > 0 ? `{ ${props.join(", ")} }` : "object";
  }
  // no type keyword (slice stubs): any value
  return "any";
}

/** Render a wrapper-union definition (each branch a single-key object). */
function renderWrapperUnion(name: string, def: Node, description: string): string[] {
  const lines = [`${name} =${description}`];
  for (const branch of def.anyOf as Node[]) {
    const properties = (branch.properties ?? {}) as Record<string, Node>;
    const entries = Object.entries(properties);
    if (entries.length === 1) {
      const [key, value] = entries[0]!;
      lines.push(`  | { ${key}: ${renderType(value)} }${trailingComment(branch)}`);
    } else {
      lines.push(`  | ${renderType(branch)}`);
    }
  }
  return lines;
}

/** Render one definition as type-definition lines. */
function renderDefinition(name: string, def: Node): string[] {
  const stubbed = typeof def.title === "string" && def.title.endsWith(" (stub)");
  if (stubbed)
    return [`${name} // stub — office-open-schema-lookup { definitions: ["${name}"] } expands it`];

  const description = headerComment(def);
  if (Array.isArray(def.enum)) {
    return [`${name} = ${def.enum.map((v) => JSON.stringify(v)).join(" | ")}${description}`];
  }
  if (Array.isArray(def.anyOf ?? def.oneOf)) {
    return renderWrapperUnion(name, def, description);
  }
  if (def.type === "object" && def.properties) {
    const required = new Set(Array.isArray(def.required) ? (def.required as string[]) : []);
    const lines = [`${name} {${description}`];
    for (const [key, value] of Object.entries(def.properties as Record<string, Node>)) {
      lines.push(
        `  ${key}${required.has(key) ? "" : "?"}: ${renderType(value)},${trailingComment(value)}`,
      );
    }
    lines.push("}");
    return lines;
  }
  return [`${name}: ${renderType(def)}${description}`];
}

/**
 * Render a slice's definitions as a type-definition document for the model.
 * Field order follows the source schema (declaration order), definitions
 * follow the slice's insertion order (requested definitions first).
 */
export function renderSliceTypeText(
  format: DocumentType,
  requested: readonly string[],
  slice: JsonSchema,
): string {
  const header =
    `office-open ${format} type definitions for: ${requested.join(", ")}. ` +
    `Lines marked "// stub" accept any value — expand them with the ` +
    `office-open-schema-lookup tool, e.g. { type: "${format}", definitions: ["StubName"] }. ` +
    '`?` marks optional fields; `"a" | "b"` enumerates allowed values.';
  const blocks: string[] = [];
  for (const [name, def] of Object.entries(slice.definitions as Record<string, Node>)) {
    blocks.push(renderDefinition(name, def).join("\n"));
  }
  return `${header}\n\n${blocks.join("\n\n")}`;
}
