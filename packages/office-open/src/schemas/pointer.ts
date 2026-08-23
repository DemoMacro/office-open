/**
 * Internal `#/definitions/…` pointer resolution.
 *
 * Definition-name refs (`#/definitions/Foo`) and the property-sharing pointers
 * the schema generator emits (`#/definitions/Foo/properties/bar`, replacing
 * schema-identical properties duplicated by extends deep-merge) both resolve
 * here so every consumer — slice extraction, skeleton derivation, type-text
 * rendering — walks them the same way.
 *
 * @module
 */

type Node = Record<string, unknown>;

/** Resolve one pointer segment: raw key first, then URI-decoded (~0/~1 last). */
function child(container: Node, segment: string): unknown {
  if (segment in container) return container[segment];
  try {
    return container[decodeURIComponent(segment.replace(/~1/g, "/").replace(/~0/g, "~"))];
  } catch {
    return undefined;
  }
}

/** A ref that reaches into a definition instead of naming one. */
export function isPropertyPointer(ref: string): boolean {
  return ref.startsWith("#/definitions/") && ref.slice("#/definitions/".length).includes("/");
}

/** Host definition name of a property pointer (`Foo` for `…/Foo/properties/bar`). */
export function pointerHost(ref: string): string {
  const [host] = ref.slice("#/definitions/".length).split("/");
  return decodeURIComponent(host!);
}

/** Resolve any `#/…` pointer against `root`; undefined when a segment is missing. */
export function resolvePointer(root: Node, ref: string): Node | undefined {
  if (!ref.startsWith("#/")) return undefined;
  let node: unknown = root;
  for (const segment of ref.slice(2).split("/")) {
    if (!node || typeof node !== "object") return undefined;
    const next = child(node as Node, segment);
    if (next === undefined) return undefined;
    node = next;
  }
  return node && typeof node === "object" ? (node as Node) : undefined;
}
