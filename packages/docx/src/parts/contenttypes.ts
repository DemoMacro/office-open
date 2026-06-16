/**
 * Content types descriptor — produces [Content_Types].xml.
 *
 * Reference: OPC, Content_Types.xsd
 *
 * @module
 */

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
    for (const d of opts.defaults) {
      p.push(defaultXml(d.extension, d.contentType));
    }
    for (const o of opts.overrides) {
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
    return { defaults, overrides } as unknown as ContentTypesInput;
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
];

/**
 * Ensure every media part in the package has a resolvable content-type Default.
 *
 * Round-tripped packages pass through the source's [Content_Types], which may
 * declare an uppercase extension (e.g. `JPG`) while the media is named `.jpg`.
 * OPC extension matching is case-sensitive, so the media ends up with no
 * content type and Word rejects it as unreadable content. This backfills any
 * missing defaults from the standard set, preserving the case actually used in
 * the media filenames.
 */
export function withMediaDefaults(
  input: ContentTypesInput,
  mediaFileNames: string[],
): ContentTypesInput {
  const have = new Set(input.defaults.map((d) => d.extension));
  const standard = new Map(STANDARD_DEFAULTS.map((d) => [d.extension, d.contentType]));
  const defaults = [...input.defaults];
  for (const fileName of mediaFileNames) {
    const ext = fileName.slice(fileName.lastIndexOf(".") + 1);
    if (!ext || have.has(ext)) continue;
    const contentType = standard.get(ext) ?? standard.get(ext.toLowerCase());
    if (contentType) {
      defaults.push({ extension: ext, contentType });
      have.add(ext);
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
  const haveExt = new Set(defaults.map((d) => d.extension));
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

/** Helper to build the standard DOCX content types with dynamic overrides. */
export function buildContentTypes(
  extras: {
    headerCount?: number;
    footerCount?: number;
    chartCount?: number;
    smartArtCount?: number;
    hasBibliography?: boolean;
    hasComments?: boolean;
    hasGlossary?: boolean;
    hasWebSettings?: boolean;
    altChunks?: { path: string; contentType: string }[];
    subDocs?: { path: string }[];
  } = {},
): ContentTypesInput {
  const defaults: ContentTypeDefault[] = [...STANDARD_DEFAULTS];

  const overrides: ContentTypeOverride[] = [
    {
      partName: "/word/document.xml",
      contentType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    },
    {
      partName: "/word/styles.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
    },
    {
      partName: "/docProps/core.xml",
      contentType: "application/vnd.openxmlformats-package.core-properties+xml",
    },
    {
      partName: "/docProps/custom.xml",
      contentType: "application/vnd.openxmlformats-officedocument.custom-properties+xml",
    },
    {
      partName: "/docProps/app.xml",
      contentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
    },
    {
      partName: "/word/numbering.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
    },
    {
      partName: "/word/footnotes.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
    },
    {
      partName: "/word/endnotes.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
    },
    {
      partName: "/word/settings.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
    },
    // Comments Override is conditional — declaring it without a comments.xml
    // part (or vice versa) is an OPC mismatch. See hasComments gate.
    ...(extras.hasComments
      ? [
          {
            partName: "/word/comments.xml",
            contentType:
              "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
          },
        ]
      : []),
    {
      partName: "/word/fontTable.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
    },
  ];

  for (let i = 1; i <= (extras.headerCount ?? 0); i++) {
    overrides.push({
      partName: `/word/header${i}.xml`,
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    });
  }
  for (let i = 1; i <= (extras.footerCount ?? 0); i++) {
    overrides.push({
      partName: `/word/footer${i}.xml`,
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
    });
  }
  for (let i = 1; i <= (extras.chartCount ?? 0); i++) {
    overrides.push({
      partName: `/word/charts/chart${i}.xml`,
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
    });
  }
  for (let i = 1; i <= (extras.smartArtCount ?? 0); i++) {
    overrides.push({
      partName: `/word/diagrams/data${i}.xml`,
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
    });
    overrides.push({
      partName: `/word/diagrams/layout${i}.xml`,
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
    });
    overrides.push({
      partName: `/word/diagrams/quickStyle${i}.xml`,
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
    });
    overrides.push({
      partName: `/word/diagrams/colors${i}.xml`,
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
    });
    overrides.push({
      partName: `/word/diagrams/drawing${i}.xml`,
      contentType: "application/vnd.ms-office.drawingml.diagramDrawing+xml",
    });
  }
  if (extras.hasBibliography) {
    overrides.push({
      partName: "/word/bibliography.xml",
      contentType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.bibliography+xml",
    });
  }
  if (extras.hasGlossary) {
    overrides.push({
      partName: "/word/glossary/document.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.glossary+xml",
    });
  }
  if (extras.hasWebSettings) {
    overrides.push({
      partName: "/word/webSettings.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
    });
  }
  for (const ac of extras.altChunks ?? []) {
    overrides.push({
      partName: ac.path.startsWith("/") ? ac.path : `/${ac.path}`,
      contentType: ac.contentType,
    });
  }
  for (const sd of extras.subDocs ?? []) {
    overrides.push({
      partName: sd.path.startsWith("/") ? sd.path : `/${sd.path}`,
      contentType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    });
  }

  return { defaults, overrides };
}
