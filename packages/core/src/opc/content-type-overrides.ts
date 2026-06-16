/**
 * Build [Content_Types].xml Override entries from a part registry.
 *
 * The per-package content-types builders previously hand-wrote every Override
 * branch (header/footer counts, the hasComments gate, …), which drifted from
 * the part registry and caused OPC mismatches (#357: a comments Override
 * declared with no part, and vice-versa). This helper drives Overrides from
 * {@link PackagePartRegistry} so the part list and its content-type
 * declarations share one source of truth.
 *
 * Only parts carrying a `contentType` produce an Override — parts that rely on
 * a `<Default>` extension mapping (`[Content_Types].xml`, `.rels`, the
 * raw-passed theme) are skipped. Dynamic parts whose count or path is
 * runtime-determined (altChunks, sub-documents) stay out of the registry and
 * are appended by the caller.
 *
 * @module
 */

import type { PackagePartRegistry } from "./part-registry";

export interface ContentTypeOverrideEntry {
  partName: string;
  contentType: string;
}

/** Ensure a part path carries the leading slash an Override PartName requires. */
function withLeadingSlash(partPath: string): string {
  return partPath.startsWith("/") ? partPath : `/${partPath}`;
}

/**
 * Derive [Content_Types].xml Override entries for every registry part present
 * under `facts`. The `facts` keys mirror the registry's `flag` / `countFrom`
 * tokens: a boolean for `conditional` parts, a count for `repeated` parts.
 * Order follows `registry.parts`; OPC does not mandate Override order.
 */
export function buildContentTypeOverrides(
  registry: PackagePartRegistry,
  facts: ReadonlyMap<string, boolean | number>,
): ContentTypeOverrideEntry[] {
  const overrides: ContentTypeOverrideEntry[] = [];
  for (const part of registry.parts) {
    // Parts without a contentType rely on a <Default> extension mapping.
    if (!part.contentType) continue;
    const presence = part.presence;
    if (presence.kind === "always") {
      // Singletons only — an `always` template with `${i}` would be a registry bug.
      if (!part.path.includes("${i}")) {
        overrides.push({ partName: withLeadingSlash(part.path), contentType: part.contentType });
      }
    } else if (presence.kind === "conditional") {
      if (facts.get(presence.flag)) {
        overrides.push({ partName: withLeadingSlash(part.path), contentType: part.contentType });
      }
    } else {
      const count = Number(facts.get(presence.countFrom) ?? 0);
      for (let i = 1; i <= count; i++) {
        overrides.push({
          partName: withLeadingSlash(part.path.replace("${i}", String(i))),
          contentType: part.contentType,
        });
      }
    }
  }
  return overrides;
}
