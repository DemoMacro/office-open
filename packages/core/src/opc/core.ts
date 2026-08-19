import { textOf, escapeXml } from "@office-open/xml";
import type { Element } from "@office-open/xml";

/**
 * Core document properties (docProps/core.xml).
 *
 * Shared across docx/pptx/xlsx: each format's top-level Options extends this,
 * parse emits the same shape, and patch overrides it — read/write symmetry
 * (CONTRIBUTING §Property Naming). Field names follow the OPC core-properties
 * XSD element local names (`creator` = dc:creator, … — never `author`).
 */
export interface CorePropertiesOptions {
  title?: string;
  subject?: string;
  creator?: string;
  keywords?: string;
  description?: string;
  lastModifiedBy?: string;
  revision?: number;
  lastPrinted?: string;
  /**
   * Creation timestamp (W3CDTF), round-tripped from dcterms:created. null =
   * the source core.xml carried none — emit nothing; undefined (fresh) =
   * default to now.
   */
  created?: string | null;
  /** Last modified timestamp (W3CDTF); same null/undefined semantics as {@link created}. */
  modified?: string | null;
  /** Document category, round-tripped from cp:category. */
  category?: string;
  /** Document status (e.g. "Draft"), round-tripped from cp:contentStatus. */
  contentStatus?: string;
  /** Unique document identifier, round-tripped from dc:identifier. */
  identifier?: string;
  /** Document language (RFC 3066), round-tripped from dc:language. */
  language?: string;
  /** Document version number, round-tripped from cp:version. */
  version?: string;
  /**
   * Emit the core-properties vocabulary as the default namespace — the ISO
   * strict binding, where the source root is `<coreProperties xmlns=…>` and
   * cp:-prefixed children appear prefix-less (dc:/dcterms: keep their
   * prefixes). Round-trip only.
   */
  defaultNamespace?: true;
}

const FIELD_MAP: Array<{ name: string; key: keyof CorePropertiesOptions }> = [
  { name: "dc:title", key: "title" },
  { name: "dc:subject", key: "subject" },
  { name: "dc:creator", key: "creator" },
  { name: "dc:description", key: "description" },
  { name: "cp:keywords", key: "keywords" },
  { name: "cp:lastModifiedBy", key: "lastModifiedBy" },
  { name: "cp:lastPrinted", key: "lastPrinted" },
  { name: "dcterms:created", key: "created" },
  { name: "dcterms:modified", key: "modified" },
  { name: "cp:category", key: "category" },
  { name: "cp:contentStatus", key: "contentStatus" },
  { name: "dc:identifier", key: "identifier" },
  { name: "dc:language", key: "language" },
  { name: "cp:version", key: "version" },
];

/**
 * Parse core properties from an already-parsed XML element.
 * Shared by docx/pptx/xlsx to extract Dublin Core metadata into the unified
 * {@link CorePropertiesOptions} shape.
 */
export function parseCorePropsElement(el: Element | undefined): CorePropertiesOptions {
  if (!el) return {};

  const props: CorePropertiesOptions = {};
  // ISO/strict binds the core-properties namespace as the default — a
  // prefix-less root means stringify must re-emit that form.
  if (el.name === "coreProperties") props.defaultNamespace = true;

  for (const field of FIELD_MAP) {
    // ISO/strict files bind the core-properties namespace as the DEFAULT
    // namespace, so cp:/dcterms: children appear prefix-less — match by
    // local name too (field local names are unique across the map).
    const localName = field.name.slice(field.name.indexOf(":") + 1);
    const child = el.elements?.find(
      (e) => e.type === "element" && (e.name === field.name || e.name === localName),
    );
    // Presence-based: Word writes whitespace-only text ("<dc:title>\n</dc:title>"),
    // which the XML parser reduces to an empty element — capture "" so the
    // field survives round-trip instead of being silently dropped.
    if (child) (props as Record<string, unknown>)[field.key] = textOf(child);
  }

  const revEl = el.elements?.find((e) => e.name === "cp:revision" || e.name === "revision");
  if (revEl) {
    const rev = textOf(revEl);
    if (rev) {
      const n = Number(rev);
      if (!Number.isNaN(n)) props.revision = n;
    }
  }

  // A parsed core part that carries no timestamps marks them explicitly
  // absent (null) — the emit path then omits them instead of defaulting to
  // now, which a fresh document (undefined) still does.
  if (props.created === undefined) props.created = null;
  if (props.modified === undefined) props.modified = null;

  return props;
}

/**
 * Build a cp:coreProperties XML string directly (fast path).
 *
 * Shared by pptx and xlsx to bypass the toXml() → xml() pipeline.
 * created/modified default to now when not supplied; all other fields emit
 * only when present.
 */
export function buildCorePropertiesXmlString(opts: CorePropertiesOptions): string {
  // ISO/strict round-trip: the core-properties namespace is the default, so
  // its children carry no prefix (dc:/dcterms: keep theirs).
  const cp = (name: string): string => (opts.defaultNamespace ? name : `cp:${name}`);
  const p: string[] = opts.defaultNamespace
    ? [
        '<coreProperties xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">',
      ]
    : [
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
      ];
  // Empty-string values are meaningful (element present, text empty) — only
  // undefined omits the element.
  if (opts.title !== undefined) p.push(`<dc:title>${escapeXml(opts.title)}</dc:title>`);
  if (opts.subject !== undefined) p.push(`<dc:subject>${escapeXml(opts.subject)}</dc:subject>`);
  if (opts.creator !== undefined) p.push(`<dc:creator>${escapeXml(opts.creator)}</dc:creator>`);
  if (opts.keywords !== undefined)
    p.push(`<${cp("keywords")}>${escapeXml(opts.keywords)}</${cp("keywords")}>`);
  if (opts.description !== undefined)
    p.push(`<dc:description>${escapeXml(opts.description)}</dc:description>`);
  if (opts.lastPrinted !== undefined)
    p.push(`<${cp("lastPrinted")}>${escapeXml(opts.lastPrinted)}</${cp("lastPrinted")}>`);
  if (opts.lastModifiedBy !== undefined)
    p.push(`<${cp("lastModifiedBy")}>${escapeXml(opts.lastModifiedBy)}</${cp("lastModifiedBy")}>`);
  if (opts.revision !== undefined)
    p.push(`<${cp("revision")}>${opts.revision}</${cp("revision")}>`);

  const now = new Date().toISOString();
  if (opts.created !== null)
    p.push(`<dcterms:created xsi:type="dcterms:W3CDTF">${opts.created ?? now}</dcterms:created>`);
  if (opts.modified !== null)
    p.push(
      `<dcterms:modified xsi:type="dcterms:W3CDTF">${opts.modified ?? now}</dcterms:modified>`,
    );
  // Trailing slots mirror Word's emission order (category last in real files).
  if (opts.category !== undefined)
    p.push(`<${cp("category")}>${escapeXml(opts.category)}</${cp("category")}>`);
  if (opts.contentStatus !== undefined)
    p.push(`<${cp("contentStatus")}>${escapeXml(opts.contentStatus)}</${cp("contentStatus")}>`);
  if (opts.identifier !== undefined)
    p.push(`<dc:identifier>${escapeXml(opts.identifier)}</dc:identifier>`);
  if (opts.language !== undefined) p.push(`<dc:language>${escapeXml(opts.language)}</dc:language>`);
  if (opts.version !== undefined)
    p.push(`<${cp("version")}>${escapeXml(opts.version)}</${cp("version")}>`);
  p.push(`</${cp("coreProperties")}>`);
  return p.join("");
}
