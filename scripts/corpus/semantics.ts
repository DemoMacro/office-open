import { createHash } from "node:crypto";

import { unzipSync } from "fflate";

import { OOXML_CANONICAL_PREFIXES, type ParsedArchive } from "../../packages/core/dist/index.mjs";
import { parse, type Element } from "../../packages/xml/dist/index.mjs";

export type SemanticPartKind = "xml" | "relationship" | "content-types" | "binary";

export type SemanticDiffCategory =
  | "missing-part"
  | "added-part"
  | "binary"
  | "element"
  | "attribute"
  | "text"
  | "child-order"
  | "child";

export interface SemanticPartDiff {
  path: string;
  kind: SemanticPartKind;
  category?: SemanticDiffCategory;
  xpath?: string;
  detail?: string;
}

interface CanonicalNode {
  name: string;
  attributes: Record<string, string>;
  text: string;
  children: CanonicalNode[];
}

const UNORDERED_PART_PATHS = new Set([
  "docProps/app.xml",
  "docProps/core.xml",
  "docProps/custom.xml",
]);

const IGNORED_ATTRIBUTES = new Set(["mc:Ignorable"]);

const STRICT_URI_PREFIX = "http://purl.oclc.org/ooxml/";
const TRANSITIONAL_URI_PREFIX = "http://schemas.openxmlformats.org/";

function canonicalAttributeValue(name: string, value: string): string {
  if (name === "Type" || name === "uri" || name === "Namespace") {
    return value.startsWith(STRICT_URI_PREFIX)
      ? `${TRANSITIONAL_URI_PREFIX}${value.slice(STRICT_URI_PREFIX.length)}`
      : value;
  }
  return value;
}

function canonicalNode(element: Element, path = ""): CanonicalNode {
  const name = element.name ?? "";
  const childPath = `${path}/${name}`;
  const isRelationship = name === "Relationship" && path.includes("_rels/");
  const attributes = Object.fromEntries(
    Object.entries(element.attributes ?? {})
      .filter(([attributeName]) => attributeName !== "xmlns" && !attributeName.startsWith("xmlns:"))
      .filter(([attributeName]) => !(isRelationship && attributeName === "Id"))
      .filter(([attributeName]) => !IGNORED_ATTRIBUTES.has(attributeName))
      .map(([attributeName, value]) => [
        attributeName,
        canonicalAttributeValue(attributeName, String(value ?? "")),
      ])
      .sort(([left], [right]) => left.localeCompare(right)),
  );
  const text = (element.elements ?? [])
    .filter((child) => child.type === "text" || child.type === "cdata")
    .map((child) => String(child.text ?? child.cdata ?? ""))
    .join("");
  return {
    name,
    attributes,
    text: element.attributes?.["xml:space"] === "preserve" ? text : text.trim(),
    children: (element.elements ?? [])
      .filter((child): child is Element => child.type === "element")
      .map((child) => canonicalNode(child, childPath)),
  };
}

function sortUnorderedChildren(node: CanonicalNode, path: string): CanonicalNode {
  const childPath = `${path}/${node.name}`;
  const children = node.children.map((child) => sortUnorderedChildren(child, childPath));
  const unorderedRoot =
    UNORDERED_PART_PATHS.has(path) || path === "[Content_Types].xml" || path.endsWith(".rels");
  return {
    ...node,
    children: unorderedRoot
      ? [...children].sort((left, right) =>
          childFingerprint(left).localeCompare(childFingerprint(right)),
        )
      : children,
  };
}

function childFingerprint(node: CanonicalNode): string {
  const hash = createHash("sha256");
  hash.update(`${node.name}\0${JSON.stringify(node.attributes)}\0${JSON.stringify(node.text)}`);
  for (const child of node.children) hash.update(childFingerprint(child));
  return hash.digest("hex");
}

function canonicalAttributes(element: Element): string {
  return Object.entries(element.attributes ?? {})
    .filter(([name]) => name !== "xmlns" && !name.startsWith("xmlns:"))
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([name, value]) => `${name}\0${value ?? ""}`)
    .join("\u0001");
}

function canonicalText(element: Element): string {
  const text = (element.elements ?? [])
    .filter((child) => child.type === "text" || child.type === "cdata")
    .map((child) => String(child.text ?? child.cdata ?? ""))
    .join("");
  return element.attributes?.["xml:space"] === "preserve" ? text : text.trim();
}

function decodeXmlBytes(bytes: Uint8Array): string {
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf)
    return new TextDecoder("utf-8").decode(bytes.subarray(3));
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe)
    return new TextDecoder("utf-16le").decode(bytes.subarray(2));
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff)
    return new TextDecoder("utf-16be").decode(bytes.subarray(2));
  const ascii = Buffer.from(bytes.subarray(0, 120)).toString("latin1");
  const declaration = /^<\?xml\s+[^>]*encoding=["']([^"']+)["']/i.exec(ascii);
  if (declaration) {
    const encoding = declaration[1]!.toLowerCase();
    if (encoding === "utf-16" || encoding === "utf16") {
      if (bytes[0] === 0 && bytes[1] !== 0) return new TextDecoder("utf-16be").decode(bytes);
      return new TextDecoder("utf-16le").decode(bytes);
    }
  }
  return new TextDecoder("utf-8").decode(bytes);
}

export function canonicalXmlDigest(element: Element | undefined): string {
  if (!element || element.type !== "element")
    return createHash("sha256").update("empty").digest("hex");
  const hash = createHash("sha256");
  const visit = (node: Element): void => {
    hash.update(`<${node.name ?? ""}|${canonicalAttributes(node)}|${canonicalText(node)}>`);
    for (const child of node.elements ?? []) {
      if (child.type === "element") visit(child);
    }
  };
  visit(element);
  return hash.digest("hex");
}

export function canonicalXmlNodes(
  element: Element | undefined,
  partPath = "",
): CanonicalNode | undefined {
  const root = element && element.type === "element" ? canonicalNode(element, partPath) : undefined;
  return root && sortUnorderedChildren(root, partPath);
}

function attributeKey(node: CanonicalNode): string {
  return Object.entries(node.attributes)
    .map(([name, value]) => `${name}=${JSON.stringify(value)}`)
    .sort()
    .join(",");
}

function childKey(node: CanonicalNode): string {
  return childFingerprint(node);
}

function compareNodes(
  path: string,
  source: CanonicalNode | undefined,
  output: CanonicalNode | undefined,
): SemanticPartDiff["detail"][] {
  if (!source && !output) return [];
  const location = path || "/";
  if (!source) return [{ category: "added-part", xpath: location, detail: "output-only element" }];
  if (!output) return [{ category: "element", xpath: location, detail: "source-only element" }];
  const diffs: NonNullable<SemanticPartDiff["detail"]>[] = [];
  if (source.name !== output.name) {
    diffs.push({
      category: "element",
      xpath: location,
      detail: `${source.name} -> ${output.name}`,
    });
  }
  for (const [name, value] of Object.entries(source.attributes)) {
    const next = output.attributes[name];
    if (next === undefined) {
      diffs.push({ category: "attribute", xpath: `${location}/@${name}`, detail: "source-only" });
    } else if (next !== value) {
      diffs.push({
        category: "attribute",
        xpath: `${location}/@${name}`,
        detail: "value mismatch",
      });
    }
  }
  for (const name of Object.keys(output.attributes)) {
    if (!(name in source.attributes)) {
      diffs.push({
        category: "attribute",
        xpath: `${location}/@${name}`,
        detail: "output-only default",
      });
    }
  }
  if (source.text !== output.text) {
    diffs.push({ category: "text", xpath: location, detail: "text mismatch" });
  }
  const sourceChildren = source.children.map((child) => ({ key: childKey(child), child }));
  const outputChildren = output.children.map((child) => ({ key: childKey(child), child }));
  const sourceCounts = new Map(sourceChildren.map(({ key }) => [key, 0]));
  for (const { key } of sourceChildren) sourceCounts.set(key, (sourceCounts.get(key) ?? 0) + 1);
  const outputCounts = new Map(outputChildren.map(({ key }) => [key, 0]));
  for (const { key } of outputChildren) outputCounts.set(key, (outputCounts.get(key) ?? 0) + 1);
  const sourceOnlyKeys = [...sourceCounts.entries()].filter(
    ([key, count]) => count > (outputCounts.get(key) ?? 0),
  );
  const outputOnlyKeys = [...outputCounts.entries()].filter(
    ([key, count]) => count > (sourceCounts.get(key) ?? 0),
  );
  if (sourceOnlyKeys.length || outputOnlyKeys.length) {
    for (const [key] of sourceOnlyKeys) {
      const child = sourceChildren.find(({ key: childKey }) => childKey === key)!.child;
      diffs.push({
        category: "child",
        xpath: `${location}/${child.name}`,
        detail: "source-only child",
      });
    }
    for (const [key] of outputOnlyKeys) {
      const child = outputChildren.find(({ key: childKey }) => childKey === key)!.child;
      diffs.push({
        category: "child",
        xpath: `${location}/${child.name}`,
        detail: "output-only child",
      });
    }
  } else if (
    sourceChildren.some(
      ({ child }, index) => childKey(child) !== childKey(outputChildren[index]!.child),
    )
  ) {
    diffs.push({ category: "child-order", xpath: location, detail: "children reordered" });
  }
  if (diffs.length) return diffs;
  for (let index = 0; index < source.children.length; index++) {
    const childDiffs = compareNodes(
      `${location}/${source.children[index]!.name}[${index + 1}]`,
      source.children[index],
      output.children[index],
    );
    if (childDiffs.length) return childDiffs;
  }
  return [];
}

export function explainSemanticPartDiff(
  partPath: string,
  source: Uint8Array,
  output: Uint8Array | undefined,
): SemanticPartDiff[] {
  const kind = semanticPartKind(partPath);
  if (!output)
    return [{ path: partPath, kind, category: "missing-part", detail: "output missing" }];
  if (kind === "binary") {
    const sourceHash = createHash("sha256").update(source).digest("hex");
    const outputHash = createHash("sha256").update(output).digest("hex");
    return sourceHash === outputHash
      ? []
      : [{ path: partPath, kind, category: "binary", detail: "bytes differ" }];
  }
  const sourceXml = parseCanonicalXml(decodeXmlBytes(source));
  const outputXml = parseCanonicalXml(decodeXmlBytes(output));
  const sourceNode = canonicalXmlNodes(sourceXml, partPath);
  const outputNode = canonicalXmlNodes(outputXml, partPath);
  const details = compareNodes(partPath, sourceNode, outputNode);
  return details.map((detail) => ({ path: partPath, kind, ...detail }));
}

function sortedElementXml(
  element: Element | undefined,
  childKey: (child: Element) => string,
): string {
  if (!element || element.type !== "element") return "empty";
  const children = (element.elements ?? [])
    .filter((child) => child.type === "element")
    .sort((left, right) => childKey(left).localeCompare(childKey(right)))
    .map((child) => `<${child.name ?? ""}|${canonicalAttributes(child)}|${canonicalText(child)}>`);
  return `${element.name ?? ""}|${canonicalAttributes(element)}|${canonicalText(element)}|${children.join("")}`;
}

export function canonicalRelationshipsDigest(element: Element | undefined): string {
  return createHash("sha256")
    .update(
      sortedElementXml(element, (child) =>
        ["Id", "Type", "Target", "TargetMode"]
          .map((name) => String(child.attributes?.[name] ?? ""))
          .join("\0"),
      ),
    )
    .digest("hex");
}

export function canonicalContentTypesDigest(element: Element | undefined): string {
  return createHash("sha256")
    .update(
      sortedElementXml(element, (child) =>
        child.name === "Default"
          ? `Default\0${String(child.attributes?.Extension ?? "")}`
          : `Override\0${String(child.attributes?.PartName ?? "")}`,
      ),
    )
    .digest("hex");
}

export function parseCanonicalXml(xml: string): Element | undefined {
  return parse(xml, {
    ignoreDeclaration: true,
    normalizeNamespaces: OOXML_CANONICAL_PREFIXES,
  }).elements?.find((candidate): candidate is Element => candidate.type === "element");
}

export function semanticPartKind(path: string): SemanticPartKind {
  if (path === "[Content_Types].xml") return "content-types";
  if (path.endsWith(".rels")) return "relationship";
  if (path.toLowerCase().endsWith(".xml")) return "xml";
  return "binary";
}

export function semanticPartDiff(
  path: string,
  source: Uint8Array,
  output: Uint8Array | undefined,
): SemanticPartDiff | undefined {
  const kind = semanticPartKind(path);
  if (kind === "binary") {
    const sourceHash = createHash("sha256").update(source).digest("hex");
    const outputHash = output
      ? createHash("sha256").update(output).digest("hex")
      : createHash("sha256").digest("hex");
    return sourceHash === outputHash ? undefined : { path, kind };
  }
  return explainSemanticPartDiff(path, source, output)[0];
}

export function archiveSemanticDiffDetails(
  source: Uint8Array,
  output: Uint8Array,
): SemanticPartDiff[] {
  const sourceArchive = unzipSync(source);
  const outputArchive = unzipSync(output);
  const paths = new Set([...Object.keys(sourceArchive), ...Object.keys(outputArchive)]);
  const diffs: SemanticPartDiff[] = [];
  for (const path of [...paths].sort()) {
    if (path.endsWith("/")) continue;
    diffs.push(
      ...explainSemanticPartDiff(
        path,
        sourceArchive[path] ?? new Uint8Array(),
        outputArchive[path],
      ),
    );
  }
  return diffs;
}

export function archiveSemanticDiffs(source: Uint8Array, output: Uint8Array): SemanticPartDiff[] {
  const sourceArchive = unzipSync(source);
  const outputArchive = unzipSync(output);
  const paths = new Set([...Object.keys(sourceArchive), ...Object.keys(outputArchive)]);
  const diffs: SemanticPartDiff[] = [];
  for (const path of [...paths].sort()) {
    if (path.endsWith("/")) continue;
    const diff = semanticPartDiff(
      path,
      sourceArchive[path] ?? new Uint8Array(),
      outputArchive[path],
    );
    if (diff) diffs.push(diff);
  }
  return diffs;
}

/** Transitional baseline comparison while strict semantic repair is in progress. */
export function archiveTagDiffs(source: Uint8Array, output: Uint8Array): string[] {
  const sourceArchive = unzipSync(source);
  const outputArchive = unzipSync(output);
  const decoder = new TextDecoder();
  const tagCounts = (xml: string): Map<string, number> => {
    const counts = new Map<string, number>();
    for (const match of xml.matchAll(/<(?:[\w-]+:)?([\w-]+)[ >/]/g)) {
      counts.set(match[1]!, (counts.get(match[1]!) ?? 0) + 1);
    }
    return counts;
  };
  const differs = (sourceXml: string, outputXml: string): boolean => {
    const sourceCounts = tagCounts(sourceXml);
    const outputCounts = tagCounts(outputXml);
    for (const [name, count] of sourceCounts) if (outputCounts.get(name) !== count) return true;
    for (const [name, count] of outputCounts) if (sourceCounts.get(name) !== count) return true;
    return false;
  };
  const diffs: string[] = [];
  for (const path of Object.keys(sourceArchive)) {
    if (!path.endsWith(".xml") && !path.endsWith(".rels")) continue;
    if (path.includes("theme")) continue;
    const sourceXml = decoder.decode(sourceArchive[path]!);
    const outputXml = decoder.decode(outputArchive[path] ?? new Uint8Array());
    if (sourceXml !== outputXml && differs(sourceXml, outputXml)) diffs.push(path);
  }
  return diffs;
}

export type CorpusFailureKind =
  | "valid"
  | "encrypted"
  | "invalid-zip"
  | "invalid-cfb"
  | "truncated-xml"
  | "invalid-package";

export function classifyPackageFailure(error: unknown): Exclude<CorpusFailureKind, "valid"> {
  const message = String(error).toLowerCase();
  if (message.includes("invalid zip data") || message.includes("end of central directory"))
    return "invalid-zip";
  if (message.includes("cfb") || message.includes("compound file")) return "invalid-cfb";
  if (message.includes("unexpected end") || message.includes("unterminated"))
    return "truncated-xml";
  if (message.includes("not found")) return "invalid-package";
  return "invalid-package";
}

export function primaryPackagePart(archive: Pick<ParsedArchive, "get">, fallback: string): string {
  const relationships = archive.get("_rels/.rels");
  for (const relationship of relationships?.elements ?? []) {
    if (relationship.name !== "Relationship") continue;
    const type = String(relationship.attributes?.Type ?? "");
    if (!type.endsWith("/officeDocument")) continue;
    if (relationship.attributes?.TargetMode === "External") continue;
    const target = String(relationship.attributes?.Target ?? "");
    if (!target) continue;
    const resolved = target.startsWith("/") ? target.slice(1) : target.replaceAll("\\", "/");
    if (archive.get(resolved)) return resolved;
  }
  if (archive.get(fallback)) return fallback;
  throw new Error(`${fallback} package primary part not found`);
}
