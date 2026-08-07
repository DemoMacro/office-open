/**
 * Content types descriptor — produces [Content_Types].xml.
 *
 * Reference: OPC, Content_Types.xsd
 *
 * @module
 */

import { buildContentTypeOverrides, DOCX_PARTS } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";

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

function defaultXml(ext: string, ct: string): string {
  return `<Default Extension="${ext}" ContentType="${ct}"/>`;
}

function overrideXml(partName: string, ct: string): string {
  return `<Override PartName="${partName}" ContentType="${ct}"/>`;
}

export const contentTypesDesc: CustomDescriptor<ContentTypesInput> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const p: string[] = [
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    ];
    // OPC Default Extension and Override PartName matching are case-insensitive
    // (ECMA-376-2 §10.1.2/§10.1.4): a duplicate differing only in case makes
    // Word reject the package as unreadable content. Dedup case-insensitively
    // as a last line of defense so no input path can emit a colliding pair.
    const seenDefault = new Set<string>();
    for (const d of opts.defaults) {
      const key = d.extension.toLowerCase();
      if (seenDefault.has(key)) continue;
      seenDefault.add(key);
      p.push(defaultXml(d.extension, d.contentType));
    }
    const seenOverride = new Set<string>();
    for (const o of opts.overrides) {
      const key = o.partName.toLowerCase();
      if (seenOverride.has(key)) continue;
      seenOverride.add(key);
      p.push(overrideXml(o.partName, o.contentType));
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

/** Standard extension → content-type defaults covering every media type emitted. */
const STANDARD_DEFAULTS: ContentTypeDefault[] = [
  { extension: "png", contentType: "image/png" },
  { extension: "jpeg", contentType: "image/jpeg" },
  { extension: "jpg", contentType: "image/jpeg" },
  { extension: "bmp", contentType: "image/bmp" },
  { extension: "gif", contentType: "image/gif" },
  { extension: "tif", contentType: "image/tiff" },
  { extension: "tiff", contentType: "image/tiff" },
  { extension: "emf", contentType: "image/x-emf" },
  { extension: "wmf", contentType: "image/x-wmf" },
  { extension: "ico", contentType: "image/x-icon" },
  { extension: "svg", contentType: "image/svg+xml" },
  { extension: "rels", contentType: "application/vnd.openxmlformats-package.relationships+xml" },
  { extension: "xml", contentType: "application/xml" },
  {
    extension: "odttf",
    contentType: "application/vnd.openxmlformats-officedocument.obfuscatedFont",
  },
  {
    extension: "bin",
    contentType: "application/vnd.openxmlformats-officedocument.oleObject",
  },
  {
    extension: "xlsx",
    contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  },
];

/**
 * Ensure every media part in the package has a resolvable content-type Default.
 *
 * Round-tripped packages pass through the source's [Content_Types], which may
 * declare an uppercase extension (e.g. `JPG`) while the media is named `.jpg`.
 * OPC extension matching is **case-insensitive** (ECMA-376-2 §10.1.2), so `JPG`
 * and `jpg` are one key — emitting both yields a duplicate default that Word
 * rejects as unreadable content. Match existing defaults case-insensitively so
 * we never add a colliding extension; the media then resolves through the
 * source's already-declared default.
 */
export function withMediaDefaults(
  input: ContentTypesInput,
  mediaFileNames: string[],
): ContentTypesInput {
  const have = new Set(input.defaults.map((d) => d.extension.toLowerCase()));
  const standard = new Map(STANDARD_DEFAULTS.map((d) => [d.extension, d.contentType]));
  const defaults = [...input.defaults];
  for (const fileName of mediaFileNames) {
    const ext = fileName.slice(fileName.lastIndexOf(".") + 1);
    const key = ext.toLowerCase();
    if (!ext || have.has(key)) continue;
    const contentType = standard.get(key);
    if (contentType) {
      defaults.push({ extension: ext, contentType });
      have.add(key);
    }
  }
  return { defaults, overrides: input.overrides };
}

/** Default content types for altChunk part extensions. */
const ALTCHUNK_DEFAULTS: Record<string, string> = {
  html: "text/html",
  rtf: "application/rtf",
  txt: "text/plain",
};

/**
 * Realign [Content_Types] Overrides for altChunk parts on a round-tripped
 * package. The pass-through content types carry Override PartNames from the
 * source, but the compiler regenerates altChunk part paths (uniqueId), so the
 * Override and the written part drift apart (O5/O6). This drops the stale
 * afchunk Overrides and appends ones matching the freshly generated paths,
 * backfilling the extension Default so the part is resolvable either way.
 */
export function withAltChunkOverrides(
  input: ContentTypesInput,
  altChunks: readonly { path: string; contentType: string }[],
): ContentTypesInput {
  const defaults = [...input.defaults];
  const haveExt = new Set(defaults.map((d) => d.extension.toLowerCase()));
  for (const ac of altChunks) {
    const ext = (ac.path.split(".").pop() ?? "").toLowerCase();
    if (ext && !haveExt.has(ext) && ALTCHUNK_DEFAULTS[ext]) {
      defaults.push({ extension: ext, contentType: ALTCHUNK_DEFAULTS[ext] });
      haveExt.add(ext);
    }
  }
  const overrides = input.overrides.filter((o) => !o.partName.startsWith("/word/afchunks/"));
  for (const ac of altChunks) {
    const partName = ac.path.startsWith("/") ? ac.path : `/${ac.path}`;
    overrides.push({ partName, contentType: ac.contentType });
  }
  return { defaults, overrides };
}

const CUSTOM_PROPERTIES_PART_NAME = "/docProps/custom.xml";
const CUSTOM_PROPERTIES_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.custom-properties+xml";

/**
 * Ensure the docProps/custom.xml Override exists on a round-tripped package.
 * The custom-properties part is always written (even when empty), but a source
 * [Content_Types] may omit its Override — the part then resolves only through
 * the generic `application/xml` Default, which Word rejects as unreadable
 * content (OPC O5). Mirror the unconditional part emission by forcing the
 * Override whenever it is missing.
 */
export function ensureCustomPropertiesOverride(input: ContentTypesInput): ContentTypesInput {
  if (input.overrides.some((o) => o.partName.toLowerCase() === CUSTOM_PROPERTIES_PART_NAME)) {
    return input;
  }
  return {
    ...input,
    overrides: [
      ...input.overrides,
      { partName: CUSTOM_PROPERTIES_PART_NAME, contentType: CUSTOM_PROPERTIES_CONTENT_TYPE },
    ],
  };
}

/**
 * Build [Content_Types].xml for a fresh DOCX compile, driven by the part
 * registry. Static parts (document/styles/…/comments/headers/…) come from
 * {@link buildContentTypeOverrides} over {@link DOCX_PARTS}; dynamic parts whose
 * path or count is runtime-determined (altChunks, sub-documents) are appended
 * here — they are not enumerable in the registry.
 *
 * `facts` keys mirror the registry's `flag` / `countFrom` tokens. The Override
 * set (order-independent) matches the former hand-written builder, so the OPC
 * consistency validator stays green.
 */
export function buildContentTypesFromRegistry(
  facts: ReadonlyMap<string, boolean | number>,
  dynamic: {
    altChunks?: ReadonlyArray<{ path: string; contentType: string }>;
    subDocs?: ReadonlyArray<{ path: string }>;
  } = {},
): ContentTypesInput {
  const overrides: ContentTypeOverride[] = buildContentTypeOverrides(DOCX_PARTS, facts);
  for (const ac of dynamic.altChunks ?? []) {
    overrides.push({
      partName: ac.path.startsWith("/") ? ac.path : `/${ac.path}`,
      contentType: ac.contentType,
    });
  }
  for (const sd of dynamic.subDocs ?? []) {
    overrides.push({
      partName: sd.path.startsWith("/") ? sd.path : `/${sd.path}`,
      contentType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    });
  }
  // Minimal defaults: rels/xml are present in every package; altChunk
  // extensions only when altChunks exist. Image/font/embedding extensions are
  // backfilled by withMediaDefaults from the actual part file names, so we no
  // longer declare every possible media type unconditionally — Word strips the
  // unused ones when repairing and that pollutes the "removed unsupported
  // content" diff.
  const defaults: ContentTypeDefault[] = [
    { extension: "rels", contentType: "application/vnd.openxmlformats-package.relationships+xml" },
    { extension: "xml", contentType: "application/xml" },
  ];
  const haveExt = new Set(["rels", "xml"]);
  for (const ac of dynamic.altChunks ?? []) {
    const ext = (ac.path.split(".").pop() ?? "").toLowerCase();
    if (ext && !haveExt.has(ext) && ALTCHUNK_DEFAULTS[ext]) {
      defaults.push({ extension: ext, contentType: ALTCHUNK_DEFAULTS[ext] });
      haveExt.add(ext);
    }
  }
  return { defaults, overrides };
}
