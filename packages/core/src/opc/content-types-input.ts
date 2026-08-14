/**
 * Files-driven [Content_Types].xml derivation (part-centric).
 *
 * Walks the actual part paths written to the package and derives the
 * Default/Override entries from a path→contentType resolver plus an
 * extension→MIME map. The part set is the single source of truth, so the
 * content-type declarations cannot drift from what is written — the class of
 * bug that registry/facts builders exist to patch (#357). Sparse or index-based
 * part naming (e.g. pptx comments keyed by slide index rather than a dense
 * sequence) is handled naturally: every written file is declared regardless of
 * its naming scheme, because emission follows the files, not a count.
 *
 * Reference: ECMA-376 Part 2 §10; peer libraries (python-pptx, openpyxl) derive
 * [Content_Types] from the actual part set the same way.
 *
 * @module
 */

import type { CustomDescriptor } from "../descriptor";
import type { PackagePartRegistry } from "./part-registry";

// ── Types ──

export interface ContentTypeDefault {
  extension: string;
  contentType: string;
}

export interface ContentTypeOverride {
  partName: string;
  contentType: string;
}

export interface ContentTypesInput {
  defaults: ContentTypeDefault[];
  overrides: ContentTypeOverride[];
}

// ── Serializer ──

export const contentTypesDesc: CustomDescriptor<ContentTypesInput> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const p: string[] = [
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    ];
    // OPC Default Extension and Override PartName matching are case-insensitive
    // (ECMA-376-2 §10.1.2/§10.1.4): a duplicate differing only in case makes
    // Word reject the package as unreadable content. Dedup case-insensitively
    // so no input path can emit a colliding pair.
    const seenDefault = new Set<string>();
    for (const d of opts.defaults) {
      const key = d.extension.toLowerCase();
      if (seenDefault.has(key)) continue;
      seenDefault.add(key);
      p.push(`<Default Extension="${d.extension}" ContentType="${d.contentType}"/>`);
    }
    const seenOverride = new Set<string>();
    for (const o of opts.overrides) {
      const key = o.partName.toLowerCase();
      if (seenOverride.has(key)) continue;
      seenOverride.add(key);
      p.push(`<Override PartName="${o.partName}" ContentType="${o.contentType}"/>`);
    }
    p.push("</Types>");
    return p.join("");
  },

  parse(el, _ctx) {
    const defaults: ContentTypeDefault[] = [];
    const overrides: ContentTypeOverride[] = [];
    for (const child of el.elements ?? []) {
      if (child.name === "Default") {
        const ext = child.attributes?.["Extension"];
        const ct = child.attributes?.["ContentType"];
        if (ext && ct) defaults.push({ extension: String(ext), contentType: String(ct) });
      } else if (child.name === "Override") {
        const pn = child.attributes?.["PartName"];
        const ct = child.attributes?.["ContentType"];
        if (pn && ct) overrides.push({ partName: String(pn), contentType: String(ct) });
      }
    }
    return { defaults, overrides };
  },
};

// ── Path → contentType resolver ──

export type PartContentTypeResolver = (partPath: string) => string | undefined;

/** Convert a registry part path (with `${i}` placeholders) to an anchored RegExp. */
function patternToRegExp(path: string): RegExp {
  const segments = path.split("${i}");
  const escaped = segments.map((s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
  return new RegExp("^" + escaped.join("\\d+") + "$");
}

/**
 * Build a path→contentType resolver from a part registry. Only the path pattern
 * and contentType are consulted; the presence/flag/countFrom fields are
 * irrelevant here because emission is driven by the files actually written, not
 * by a separate facts map. Any path the resolver does not recognize falls
 * through to extension-based Default handling (media, embeddings, .rels).
 */
export function resolverFromRegistry(registry: PackagePartRegistry): PartContentTypeResolver {
  const patterns: Array<{ re: RegExp; contentType: string }> = [];
  for (const part of registry.parts) {
    if (!part.contentType) continue;
    patterns.push({ re: patternToRegExp(part.path), contentType: part.contentType });
  }
  return (partPath: string) => patterns.find((p) => p.re.test(partPath))?.contentType;
}

// ── Derive ──

const BASE_DEFAULTS: ContentTypeDefault[] = [
  { extension: "rels", contentType: "application/vnd.openxmlformats-package.relationships+xml" },
  { extension: "xml", contentType: "application/xml" },
];

export interface DeriveContentTypesOptions {
  /** Resolve an OOXML part path to its Override content type. Return undefined
   * for non-part files (media/embeddings/.rels) resolved via Default. */
  resolve: PartContentTypeResolver;
  /** Lowercase extension → MIME for media/embedding Default entries. */
  mediaContentTypes: Readonly<Record<string, string>>;
  /** Always-present Default entries. Defaults to rels + xml. */
  baseDefaults?: ContentTypeDefault[];
  /** Per-file Overrides for parts whose content type is data-driven and not
   * path-determinable — docx altChunks carry a caller-supplied MIME, and
   * sub-documents live at an arbitrary path. Each path is normalized to a
   * leading slash and takes precedence over a resolver match for the same
   * path. */
  overrides?: ReadonlyArray<{ path: string; contentType: string }>;
}

function extensionOf(path: string): string | undefined {
  const slash = Math.max(path.lastIndexOf("/"), path.lastIndexOf("\\"));
  const dot = path.slice(slash + 1).lastIndexOf(".");
  return dot > 0 ? path.slice(slash + 1 + dot + 1) : undefined;
}

function withLeadingSlash(path: string): string {
  return path.startsWith("/") ? path : `/${path}`;
}

/**
 * Derive [Content_Types].xml entries from the actual part paths in the package.
 * Each path that the resolver maps to a content type becomes an Override; every
 * other file contributes a Default keyed by its extension (declared once).
 */
export function deriveContentTypes(
  filePaths: Iterable<string>,
  options: DeriveContentTypesOptions,
): ContentTypesInput {
  const defaults: ContentTypeDefault[] = [...(options.baseDefaults ?? BASE_DEFAULTS)];
  const seenExt = new Set(defaults.map((d) => d.extension.toLowerCase()));
  // Keyed by lowercased PartName so explicit `overrides` entries (altChunks,
  // sub-documents) take precedence over a resolver match for the same path.
  const overrideMap = new Map<string, ContentTypeOverride>();

  for (const path of filePaths) {
    if (path === "[Content_Types].xml") continue;
    const partType = options.resolve(path);
    if (partType) {
      const partName = withLeadingSlash(path);
      overrideMap.set(partName.toLowerCase(), { partName, contentType: partType });
      continue;
    }
    const ext = extensionOf(path);
    if (!ext) continue;
    const key = ext.toLowerCase();
    if (seenExt.has(key)) continue;
    const mime = options.mediaContentTypes[key];
    if (mime) {
      defaults.push({ extension: ext, contentType: mime });
      seenExt.add(key);
    }
  }
  for (const extra of options.overrides ?? []) {
    const partName = withLeadingSlash(extra.path);
    overrideMap.set(partName.toLowerCase(), { partName, contentType: extra.contentType });
  }
  return { defaults, overrides: [...overrideMap.values()] };
}

/**
 * Extension → MIME for image media, shared by every format package.
 *
 * The three format compilers each declare a package-scoped media table made
 * of this image block plus format-specific extras (docx: odttf/bin fonts and
 * OLE; pptx: video/audio; xlsx: vml). Keeping the image half in one place
 * means a new image format lands everywhere at once.
 */
export const IMAGE_MEDIA_CONTENT_TYPES: Record<string, string> = {
  png: "image/png",
  jpeg: "image/jpeg",
  jpg: "image/jpeg",
  bmp: "image/bmp",
  gif: "image/gif",
  tif: "image/tiff",
  tiff: "image/tiff",
  emf: "image/x-emf",
  wmf: "image/x-wmf",
  ico: "image/x-icon",
  svg: "image/svg+xml",
};
