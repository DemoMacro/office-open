/**
 * Package-wide passthrough collection (SDK ExtendedPart analogue).
 *
 * The parse side of the "maximum parse" contract: every part the model did
 * NOT absorb is carried verbatim instead of dropped. Callers list the parts
 * their compiler rebuilds; everything else — unknown extensions, companion
 * .rels, media referenced only by passthrough parts — survives round-trip
 * with bytes and content-type declarations intact.
 *
 * Rebuilt parts whose source .rels referenced a passthrough part are captured
 * as relationship entries so the compiler can re-emit them verbatim (fresh
 * rId, unchanged target — passthrough paths never move).
 *
 * @module
 */

import { attr } from "@office-open/xml";

import {
  contentTypesDesc,
  type ContentTypeDefault,
  type ContentTypeOverride,
} from "./content-types-input";
import type { ParsedArchive } from "./parser";
import { partPathToRelsPath, resolveRelationshipTarget } from "./relationships";

// ── Types ──

export interface PassthroughPart {
  path: string;
  data: Uint8Array;
  /**
   * Content type borrowed from the source [Content_Types].xml: the part's
   * Override value, or its extension's Default. Undefined when the source
   * declared neither — carried through so the compiler can omit the entry
   * exactly like the source did.
   */
  contentType?: string;
}

export interface PassthroughRelationship {
  /** The rebuilt part whose source .rels referenced a passthrough part. */
  source: string;
  relationshipType: string;
  /** Target exactly as written in the source .rels (relative form). */
  target: string;
}

export interface PassthroughResult {
  parts: PassthroughPart[];
  relationships: PassthroughRelationship[];
}

// ── Collection ──

/** Package-level files the packer always regenerates. */
const ALWAYS_REBUILT = new Set(["[Content_Types].xml", "_rels/.rels"]);

function extensionOf(path: string): string | undefined {
  const slash = path.lastIndexOf("/");
  const dot = path.slice(slash + 1).lastIndexOf(".");
  return dot > 0 ? path.slice(slash + 1 + dot + 1) : undefined;
}

/**
 * Collect passthrough parts and rebuilt→passthrough relationships.
 *
 * @param archive - Parsed source package.
 * @param rebuiltPaths - Parts (including their .rels) the compiler re-emits
 *   from the model. Listing a part here without its .rels leaves the .rels
 *   to pass through — used by docx glossary, where the document part is
 *   rebuilt but its companion rels travel verbatim.
 */
export function collectPassthroughParts(
  archive: ParsedArchive,
  rebuiltPaths: Iterable<string>,
): PassthroughResult {
  const rebuilt = new Set<string>(ALWAYS_REBUILT);
  for (const p of rebuiltPaths) rebuilt.add(p);

  // Borrowed content-type resolution: Override by part name first, then the
  // extension's Default (both case-insensitive, matching OPC matching rules).
  let defaults: ContentTypeDefault[] = [];
  let overrides: ContentTypeOverride[] = [];
  const ctEl = archive.get("[Content_Types].xml");
  if (ctEl) {
    const parsed = contentTypesDesc.parse(ctEl, undefined as never);
    defaults = parsed.defaults;
    overrides = parsed.overrides;
  }
  const overrideMap = new Map(
    overrides.map((o) => [o.partName.toLowerCase(), o.contentType] as const),
  );
  const defaultMap = new Map(
    defaults.map((d) => [d.extension.toLowerCase(), d.contentType] as const),
  );
  const contentTypeFor = (path: string): string | undefined => {
    const override = overrideMap.get(`/${path}`.toLowerCase());
    if (override) return override;
    const ext = extensionOf(path);
    return ext ? defaultMap.get(ext.toLowerCase()) : undefined;
  };

  // Everything not rebuilt passes through. The default is "keep": companion
  // .rels of passthrough parts and media referenced only through them fall
  // out of this same loop — no closure walk needed, because not-rebuilt IS
  // the kept set.
  const parts: PassthroughPart[] = [];
  const kept = new Set<string>();
  for (const path of archive.keys()) {
    if (path.endsWith("/") || rebuilt.has(path)) continue;
    const data = archive.getRaw(path);
    if (!data) continue;
    kept.add(path);
    const contentType = contentTypeFor(path);
    parts.push(contentType === undefined ? { path, data } : { path, data, contentType });
  }

  // Rebuilt parts whose source .rels point at passthrough parts: keep the
  // relationship so the rebuilt rels stay byte-equivalent to the source.
  // External targets (URLs) are not part references — the caller handles
  // those through their own model fields.
  const relationships: PassthroughRelationship[] = [];
  for (const source of rebuiltPaths) {
    const relsEl = archive.get(partPathToRelsPath(source));
    if (!relsEl) continue;
    for (const rel of relsEl.elements ?? []) {
      if (rel.name !== "Relationship") continue;
      if (attr(rel, "TargetMode") === "External") continue;
      const relationshipType = attr(rel, "Type");
      const target = attr(rel, "Target");
      if (!relationshipType || !target) continue;
      if (kept.has(resolveRelationshipTarget(source, target))) {
        relationships.push({ source, relationshipType, target });
      }
    }
  }

  return { parts, relationships };
}
