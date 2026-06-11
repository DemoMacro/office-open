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
    return { defaults, overrides } as Record<string, unknown>;
  },
};

/** Helper to build the standard DOCX content types with dynamic overrides. */
export function buildContentTypes(
  extras: {
    headerCount?: number;
    footerCount?: number;
    chartCount?: number;
    smartArtCount?: number;
    hasBibliography?: boolean;
    hasGlossary?: boolean;
    hasWebSettings?: boolean;
    altChunks?: { path: string; contentType: string }[];
    subDocs?: { path: string }[];
  } = {},
): ContentTypesInput {
  const defaults: ContentTypeDefault[] = [
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
    {
      partName: "/word/comments.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
    },
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
