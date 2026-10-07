import { textOf, escapeXml } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { ReproducibleScope } from "../util/reproducible";

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
  /** Revision number (cp:revision). null preserves a source element with
   * empty text; undefined omits the element on fresh authoring. */
  revision?: number | null;
  lastPrinted?: string;
  /**
   * Creation timestamp (W3CDTF), round-tripped from dcterms:created. null =
   * the source core.xml carried none — emit nothing; undefined (fresh) =
   * default to now.
   */
  created?: string | null;
  /** Last modified timestamp (W3CDTF); same null/undefined semantics as the created field. */
  modified?: string | null;
  /** Document category, round-tripped from cp:category. */
  category?: string;
  /** Document status (e.g. "Draft"), round-tripped from cp:contentStatus. */
  contentStatus?: string;
  /** Content type (e.g. "Document"), round-tripped from cp:contentType. */
  contentType?: string;
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
  /** Legacy Microsoft core-property root/element spelling; round-trip only. */
  legacyMicrosoft?: true;
  /** Source child order as canonical field keys; round-trip only. */
  childOrder?: CorePropertiesChild[];
}

const LEGACY_CORE_PROPERTIES_NAMESPACE =
  "http://schemas.microsoft.com/package/2005/06/metadata/core-properties";

type CorePropertiesChild =
  | "title"
  | "subject"
  | "creator"
  | "keywords"
  | "description"
  | "lastModifiedBy"
  | "lastPrinted"
  | "revision"
  | "created"
  | "modified"
  | "category"
  | "contentStatus"
  | "contentType"
  | "identifier"
  | "language"
  | "version";

const CANONICAL_CORE_PROPERTIES_ORDER: CorePropertiesChild[] = [
  "title",
  "subject",
  "creator",
  "keywords",
  "description",
  "lastPrinted",
  "lastModifiedBy",
  "revision",
  "created",
  "modified",
  "category",
  "contentStatus",
  "contentType",
  "identifier",
  "language",
  "version",
];

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
  { name: "cp:contentType", key: "contentType" },
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
  if (el.attributes?.xmlns === LEGACY_CORE_PROPERTIES_NAMESPACE) props.legacyMicrosoft = true;
  else if (el.name === "coreProperties") props.defaultNamespace = true;

  for (const field of FIELD_MAP) {
    // ISO/strict files bind the core-properties namespace as the DEFAULT
    // namespace, so cp:/dcterms: children appear prefix-less — match by
    // local name too (field local names are unique across the map).
    const localName = field.name.slice(field.name.indexOf(":") + 1);
    const normalizedLocalName = localName.toLowerCase();
    const legacyAliases =
      normalizedLocalName === "created"
        ? ["datecreated"]
        : normalizedLocalName === "modified"
          ? ["datemodified"]
          : [];
    const child = el.elements?.find(
      (e) =>
        e.type === "element" &&
        (e.name === field.name ||
          e.name === localName ||
          legacyAliases.includes(e.name?.slice(e.name.indexOf(":") + 1).toLowerCase() ?? "") ||
          e.name?.slice(e.name.indexOf(":") + 1).toLowerCase() === normalizedLocalName),
    );
    // Presence-based: Word writes whitespace-only text ("<dc:title>\n</dc:title>"),
    // which the XML parser reduces to an empty element — capture "" so the
    // field survives round-trip instead of being silently dropped.
    if (child) (props as Record<string, unknown>)[field.key] = textOf(child);
  }

  const revEl = el.elements?.find(
    (e) =>
      e.name === "cp:revision" ||
      e.name === "revision" ||
      e.name?.slice(e.name.indexOf(":") + 1).toLowerCase() === "revision",
  );
  if (revEl) {
    const rev = textOf(revEl);
    if (rev === "") {
      // Empty source element — preserve its presence without inventing a value.
      props.revision = null;
    } else if (rev) {
      const n = Number(rev);
      if (!Number.isNaN(n)) props.revision = n;
    }
  }

  // A parsed core part that carries no timestamps marks them explicitly
  // absent (null) — the emit path then omits them instead of defaulting to
  // now, which a fresh document (undefined) still does.
  if (props.created === undefined) props.created = null;
  if (props.modified === undefined) props.modified = null;

  const childOrder = el.elements
    ?.filter((child) => child.type === "element")
    .map((child) => {
      const localName = child.name?.slice(child.name.indexOf(":") + 1).toLowerCase();
      const field = FIELD_MAP.find((candidate) => {
        const candidateLocal = candidate.name.slice(candidate.name.indexOf(":") + 1);
        return (
          child.name === candidate.name ||
          child.name === candidateLocal ||
          localName === candidateLocal.toLowerCase()
        );
      });
      if (field) return field.key;
      if (localName === "revision") return "revision" as const;
      if (localName === "datecreated") return "created" as const;
      if (localName === "datemodified") return "modified" as const;
      return undefined;
    })
    .filter((key): key is CorePropertiesChild => key !== undefined);
  let hasReorderedChildren = false;
  let canonicalIndex = -1;
  for (const key of childOrder ?? []) {
    const index = CANONICAL_CORE_PROPERTIES_ORDER.indexOf(key);
    if (index === -1 || index < canonicalIndex) {
      hasReorderedChildren = true;
      break;
    }
    canonicalIndex = index;
  }
  if (childOrder && childOrder.length > 1 && hasReorderedChildren) {
    props.childOrder = childOrder;
  }

  return props;
}

/**
 * Build a cp:coreProperties XML string directly (fast path).
 *
 * Shared by pptx and xlsx to bypass the toXml() → xml() pipeline.
 * created/modified default to now when not supplied (the reproducible scope's
 * date when one is passed); all other fields emit only when present.
 */
export function buildCorePropertiesXmlString(
  opts: CorePropertiesOptions,
  reproducible?: ReproducibleScope,
): string {
  // ISO/strict round-trip: the core-properties namespace is the default, so
  // its children carry no prefix (dc:/dcterms: keep theirs).
  const cp = (name: string): string => (opts.defaultNamespace ? name : `cp:${name}`);
  const legacy = opts.legacyMicrosoft === true;
  const legacyName = (name: string): string => {
    switch (name) {
      case "keywords":
        return "Keywords";
      case "title":
        return "Title";
      case "subject":
        return "Subject";
      case "creator":
        return "Creator";
      case "description":
        return "Description";
      case "identifier":
        return "Identifier";
      case "language":
        return "Language";
      case "lastPrinted":
        return "LastPrinted";
      case "lastModifiedBy":
        return "LastModifiedBy";
      case "category":
        return "Category";
      case "contentStatus":
        return "ContentStatus";
      case "contentType":
        return "ContentType";
      case "version":
        return "Version";
      default:
        return name;
    }
  };
  const dublinCore = (name: string): string => (legacy ? legacyName(name) : `dc:${name}`);
  const coreProperty = (name: string): string => (legacy ? legacyName(name) : cp(name));
  const open = legacy
    ? [
        '<CoreProperties xmlns="http://schemas.microsoft.com/package/2005/06/metadata/core-properties">',
      ]
    : opts.defaultNamespace
      ? [
          '<coreProperties xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">',
        ]
      : [
          '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
        ];
  const now = reproducible?.date ?? new Date().toISOString();
  const entries: Array<[CorePropertiesChild, string]> = [];
  const add = (key: CorePropertiesChild, value: string | undefined, xml?: string): void => {
    if (value !== undefined && value !== null) entries.push([key, xml ?? value]);
  };
  add(
    "title",
    opts.title,
    `<${dublinCore("title")}>${escapeXml(opts.title ?? "")}</${dublinCore("title")}>`,
  );
  add(
    "subject",
    opts.subject,
    `<${dublinCore("subject")}>${escapeXml(opts.subject ?? "")}</${dublinCore("subject")}>`,
  );
  add(
    "creator",
    opts.creator,
    `<${dublinCore("creator")}>${escapeXml(opts.creator ?? "")}</${dublinCore("creator")}>`,
  );
  add(
    "keywords",
    opts.keywords,
    `<${coreProperty("keywords")}>${escapeXml(opts.keywords ?? "")}</${coreProperty("keywords")}>`,
  );
  add(
    "description",
    opts.description,
    `<${dublinCore("description")}>${escapeXml(opts.description ?? "")}</${dublinCore("description")}>`,
  );
  add(
    "lastPrinted",
    opts.lastPrinted,
    `<${coreProperty("lastPrinted")}>${escapeXml(opts.lastPrinted ?? "")}</${coreProperty("lastPrinted")}>`,
  );
  add(
    "lastModifiedBy",
    opts.lastModifiedBy,
    `<${coreProperty("lastModifiedBy")}>${escapeXml(opts.lastModifiedBy ?? "")}</${coreProperty("lastModifiedBy")}>`,
  );
  if (opts.revision !== undefined)
    entries.push([
      "revision",
      `<${coreProperty("revision")}>${opts.revision ?? ""}</${coreProperty("revision")}>`,
    ]);
  if (opts.created !== null)
    add(
      "created",
      opts.created ?? now,
      legacy
        ? `<DateCreated>${opts.created ?? now}</DateCreated>`
        : `<dcterms:created xsi:type="dcterms:W3CDTF">${opts.created ?? now}</dcterms:created>`,
    );
  if (opts.modified !== null)
    add(
      "modified",
      opts.modified ?? now,
      legacy
        ? `<DateModified>${opts.modified ?? now}</DateModified>`
        : `<dcterms:modified xsi:type="dcterms:W3CDTF">${opts.modified ?? now}</dcterms:modified>`,
    );
  add(
    "category",
    opts.category,
    `<${coreProperty("category")}>${escapeXml(opts.category ?? "")}</${coreProperty("category")}>`,
  );
  add(
    "contentStatus",
    opts.contentStatus,
    `<${coreProperty("contentStatus")}>${escapeXml(opts.contentStatus ?? "")}</${coreProperty("contentStatus")}>`,
  );
  add(
    "contentType",
    opts.contentType,
    `<${coreProperty("contentType")}>${escapeXml(opts.contentType ?? "")}</${coreProperty("contentType")}>`,
  );
  add(
    "identifier",
    opts.identifier,
    `<dc:identifier>${escapeXml(opts.identifier ?? "")}</dc:identifier>`,
  );
  add("language", opts.language, `<dc:language>${escapeXml(opts.language ?? "")}</dc:language>`);
  add(
    "version",
    opts.version,
    `<${coreProperty("version")}>${escapeXml(opts.version ?? "")}</${coreProperty("version")}>`,
  );
  const order = opts.childOrder;
  entries.sort((left, right) => {
    if (!order) return 0;
    const leftIndex = order.indexOf(left[0]);
    const rightIndex = order.indexOf(right[0]);
    return (
      (leftIndex === -1 ? order.length : leftIndex) -
      (rightIndex === -1 ? order.length : rightIndex)
    );
  });
  const p: string[] = [
    ...open,
    ...entries.map(([, xml]) => xml),
    legacy ? "</CoreProperties>" : `</${cp("coreProperties")}>`,
  ];
  return p.join("");
}
