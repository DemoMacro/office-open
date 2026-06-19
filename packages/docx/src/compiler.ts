/**
 * DOCX document compiler — pure function entry point.
 *
 * compileDocument() accepts DocumentOptions directly,
 * creates a DocxWriteContext internally, and produces a Zippable result.
 * All XML parts are produced via descriptors or serialize() —
 * no Formatter dependency.
 *
 * @module
 */

import {
  addSmartArtRelationships,
  findAndReplaceImagePlaceholders,
  formatId,
  hasPlaceholders,
  levelForMediaName,
  replaceAllPlaceholders,
  replaceNumberingPlaceholders,
} from "@office-open/core";
import type { XmlifyedFile, ZipOptions, Zippable } from "@office-open/core";
import {
  DEFAULT_DRAWING_XML,
  getColorXml,
  getLayoutXml,
  getStyleXml,
} from "@office-open/core/smartart";
import type { DocumentOptions } from "@parts/core-properties";
import { obfuscate } from "@parts/fonts/obfuscate-ttf-to-odttf";
import { HEADER_NAMESPACES, FOOTER_NAMESPACES, stringifyHeaderFooter } from "@parts/header-footer";

import { stringifyDocumentXml, stringifyBodyChild, type BodyContext } from "./body";
import { DocxWriteContext } from "./context";
import {
  corePropertiesDesc,
  customPropertiesDesc,
  appPropertiesDesc,
  contentTypesDesc,
  buildContentTypesFromRegistry,
  withAltChunkOverrides,
  withMediaDefaults,
  fontTableDesc,
  webSettingsDesc,
  commentsDesc,
  bibliographyDesc,
  settingsDesc,
  footnotesDesc,
  endnotesDesc,
  glossaryDesc,
} from "./parts";

/** Reusable TextEncoder (stateless, safe to share). */
const encoder = new TextEncoder();

/** XML declaration prepended to every OOXML part. */
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** Extended context for header/footer formatted view caching. */
type DocxContext = BodyContext & {
  headerFormattedViews?: Map<number, string>;
  footerFormattedViews?: Map<number, string>;
};

// ── Public API ──

/**
 * Compile document options into a flat file map suitable for fflate zipSync.
 *
 * This is the primary entry point for DOCX generation — accepts DocumentOptions
 * directly.
 */
export function compileDocument(
  options: DocumentOptions,
  overrides: XmlifyedFile[] = [],
  mediaLevel: number = 0,
): Zippable {
  const ctx = new DocxWriteContext(options);
  const files: Zippable = {};

  const headerFormattedViews = new Map<number, string>();
  const footerFormattedViews = new Map<number, string>();

  const xmlifiedFileMapping = xmlifyContext(ctx, headerFormattedViews, footerFormattedViews);
  const map = new Map<string, XmlifyedFile | XmlifyedFile[]>(Object.entries(xmlifiedFileMapping));

  for (const [, obj] of map) {
    if (obj === undefined) continue;
    if (Array.isArray(obj)) {
      for (const subFile of obj) {
        files[subFile.path] =
          typeof subFile.data === "string" ? encoder.encode(subFile.data) : subFile.data;
      }
    } else {
      files[obj.path] = typeof obj.data === "string" ? encoder.encode(obj.data) : obj.data;
    }
  }

  for (const subFile of overrides) {
    files[subFile.path] =
      typeof subFile.data === "string" ? encoder.encode(subFile.data) : subFile.data;
  }

  // Media files
  const mediaArray = ctx.media.array;
  for (const mediaData of mediaArray) {
    files[`word/media/${mediaData.fileName}`] = [
      mediaData.data as Uint8Array,
      { level: levelForMediaName(mediaData.fileName, mediaLevel) as ZipOptions["level"] },
    ];
    if (mediaData.type === "svg") {
      files[`word/media/${mediaData.fallback.fileName}`] = [
        mediaData.fallback.data as Uint8Array,
        {
          level: levelForMediaName(mediaData.fallback.fileName, mediaLevel) as ZipOptions["level"],
        },
      ];
    }
  }

  // Font files — only fonts carrying binary data produce a .odttf part.
  // Round-tripped fonts (rawOdttf) keep their original obfuscated bytes.
  for (const font of ctx.fontTable.fontOptionsWithKey) {
    if (font.data === undefined) continue;
    const [nameWithoutExtension] = font.name.split(".");
    const filePath = font.odttfPath ?? `word/fonts/${nameWithoutExtension}.odttf`;
    files[filePath] = font.rawOdttf ? font.data : obfuscate(font.data, font.fontKey);
  }

  // Raw passthrough parts (word/theme/*, customXml/*, …) — generate doesn't
  // rebuild these, so copy their original bytes verbatim to keep [Content_Types]
  // declarations valid and the package openable in Word.
  for (const part of ctx._options.rawParts ?? []) {
    files[part.path] = part.data;
  }

  return files;
}

// ── Internal ──

/**
 * Complete mapping of all XML files in an OOXML document package.
 */
interface XmlifyedFileMapping {
  Document: XmlifyedFile;
  Styles: XmlifyedFile;
  Properties: XmlifyedFile;
  Numbering: XmlifyedFile;
  Relationships: XmlifyedFile;
  FileRelationships: XmlifyedFile;
  Headers: XmlifyedFile[];
  Footers: XmlifyedFile[];
  HeaderRelationships: XmlifyedFile[];
  FooterRelationships: XmlifyedFile[];
  ContentTypes: XmlifyedFile;
  CustomProperties: XmlifyedFile;
  AppProperties: XmlifyedFile;
  FootNotes: XmlifyedFile;
  FootNotesRelationships?: XmlifyedFile;
  Endnotes: XmlifyedFile;
  EndnotesRelationships?: XmlifyedFile;
  Settings: XmlifyedFile;
  Comments?: XmlifyedFile;
  CommentsRelationships?: XmlifyedFile;
  FontTable?: XmlifyedFile;
  FontTableRelationships?: XmlifyedFile;
  Bibliography?: XmlifyedFile;
  Charts?: XmlifyedFile[];
  DiagramData?: XmlifyedFile[];
  DiagramLayout?: XmlifyedFile[];
  DiagramStyle?: XmlifyedFile[];
  DiagramColors?: XmlifyedFile[];
  DiagramDrawing?: XmlifyedFile[];
  AltChunks?: XmlifyedFile[];
  SubDocs?: XmlifyedFile[];
  Glossary?: XmlifyedFile;
  WebSettings?: XmlifyedFile;
}

function xmlifyContext(
  ctx: DocxWriteContext,
  headerFormattedViews: Map<number, string>,
  footerFormattedViews: Map<number, string>,
): XmlifyedFileMapping {
  const mkCtx = (viewWrapper: DocxContext["viewWrapper"] = ctx.document): DocxContext => ({
    fileData: ctx,
    file: ctx,
    viewWrapper,
    stringifyChild: stringifyBodyChild,
    addRelationship: (type: string, target: string, mode?: string) =>
      ctx.addRelationship(type, target, mode),
    addMedia: (data: Uint8Array, type: string) => ctx.addMedia(data, type),
  });

  const documentRelationshipCount = ctx.document.relationships.relationshipCount + 1;
  // Per-part media-replacement results shared between the .rels pass and the
  // body-XML pass so both use identical rId offsets. Each header/footer part
  // has its own relationship numbering (independent of the document part).
  const footerMediaResults = new Map<number, { xml: string; referenced: { fileName: string }[] }>();
  const headerMediaResults = new Map<number, { xml: string; referenced: { fileName: string }[] }>();
  const docCtx = mkCtx(ctx.document);
  const documentXmlData = XML_DECL + stringifyDocumentXml(ctx, docCtx);

  // Comments is an optional part — skip it entirely (no comments.xml, no
  // comments rels, no [Content_Types] Override) when the document carries none.
  // Emitting an empty comments.xml with a dangling relationship is the OPC
  // violation that makes Word reject the package on open.
  const hasComments = !!ctx._options.comments?.children?.length;
  const commentRelationshipCount = hasComments
    ? ctx.comments.relationships.relationshipCount + 1
    : 0;
  const commentCtx = hasComments ? mkCtx({ relationships: ctx.comments.relationships }) : null;
  const commentXmlData = commentCtx
    ? XML_DECL + commentsDesc.stringify(ctx._options.comments!, commentCtx)
    : "";

  const footnoteRelationshipCount = ctx.footNotes.relationships.relationshipCount + 1;
  const footnoteCtx = mkCtx({
    relationships: ctx.footNotes.relationships,
  });
  const footnoteXmlData =
    XML_DECL +
    (footnotesDesc.stringify(
      {
        notes: ctx.footNotes.notes,
        separator: ctx.footNotes.separator,
        continuationSeparator: ctx.footNotes.continuationSeparator,
      },
      footnoteCtx,
    ) ?? "");

  const documentMedia = findAndReplaceImagePlaceholders(
    documentXmlData,
    ctx.media.array,
    documentRelationshipCount,
  );
  const commentMedia = hasComments
    ? findAndReplaceImagePlaceholders(commentXmlData, ctx.media.array, commentRelationshipCount)
    : { xml: "", referenced: [] as { fileName: string }[] };
  const footnoteMedia = findAndReplaceImagePlaceholders(
    footnoteXmlData,
    ctx.media.array,
    footnoteRelationshipCount,
  );
  // Register footnote media relationships eagerly so the relationshipCount used
  // to gate footnotes.xml.rels reflects the final state (see FootNotesRelationships).
  for (let i = 0; i < footnoteMedia.referenced.length; i++) {
    ctx.footNotes.relationships.addRelationship(
      footnoteRelationshipCount + i,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
      `media/${footnoteMedia.referenced[i].fileName}`,
    );
  }

  return {
    AppProperties: {
      data: XML_DECL + (appPropertiesDesc.stringify(ctx._options.appProperties ?? {}, ctx) ?? ""),
      path: "docProps/app.xml",
    },
    ...(hasComments
      ? {
          Comments: {
            data: (() => {
              const xmlData =
                commentMedia.referenced.length > 0 ? commentMedia.xml : commentXmlData;
              return replaceNumberingPlaceholders(xmlData, ctx.numbering.concreteNumbering);
            })(),
            path: "word/comments.xml",
          },
          CommentsRelationships: {
            data: (() => {
              for (let i = 0; i < commentMedia.referenced.length; i++) {
                ctx.comments.relationships.addRelationship(
                  commentRelationshipCount + i,
                  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                  `media/${commentMedia.referenced[i].fileName}`,
                );
              }
              return XML_DECL + ctx.comments.relationships.serialize();
            })(),
            path: "word/_rels/comments.xml.rels",
          },
        }
      : {}),
    ContentTypes: {
      data:
        XML_DECL +
        (contentTypesDesc.stringify(
          withMediaDefaults(
            (() => {
              const altChunks = ctx.altChunks.array.map((ac) => ({
                path: `/word/${ac.path}`,
                contentType: ac.contentType ?? "application/xhtml+xml",
              }));
              // Round-trip passes the source [Content_Types] through, but the
              // compiler regenerates altChunk part paths — realign the afchunk
              // Overrides to the freshly written parts (else O5/O6).
              return ctx._options.contentTypes
                ? withAltChunkOverrides(ctx._options.contentTypes, altChunks)
                : buildContentTypesFromRegistry(
                    new Map<string, boolean | number>([
                      ["freshCompile", true],
                      ["hasComments", hasComments],
                      ["hasBibliography", !!ctx._options.bibliography],
                      ["hasGlossary", !!ctx.glossaryOptions],
                      ["hasWebSettings", !!ctx.webSettings],
                      ["headerCount", ctx.headers.length],
                      ["footerCount", ctx.footers.length],
                      ["chartCount", ctx.charts.array.length],
                      ["smartArtCount", ctx.smartArts.array.length],
                    ]),
                    {
                      altChunks,
                      subDocs: ctx.subDocs.array.map((sd) => ({ path: `/word/${sd.path}` })),
                    },
                  );
            })(),
            ctx.media.array.map((m) => m.fileName),
          ),
          ctx,
        ) ?? ""),
      path: "[Content_Types].xml",
    },
    CustomProperties: {
      data:
        XML_DECL +
        (customPropertiesDesc.stringify({ properties: ctx._options.customProperties ?? [] }, ctx) ??
          ""),
      path: "docProps/custom.xml",
    },
    Document: {
      data: (() => {
        let xmlData = documentMedia.referenced.length > 0 ? documentMedia.xml : documentXmlData;
        if (hasPlaceholders(xmlData)) {
          const mediaCount = documentMedia.referenced.length;
          const chartKeys = ctx.charts.array.map((c) => c.key);
          const smartArtKeys = ctx.smartArts.array.map((s) => s.key);
          const chartOffset = documentRelationshipCount + mediaCount;
          const smartArtOffset = chartOffset + chartKeys.length;

          // Build combined replacement entries for charts, smartart, and numbering
          const entries: Array<{ prefix?: string; key: string; value: string }> = [];
          for (let i = 0; i < chartKeys.length; i++) {
            entries.push({
              prefix: "chart:",
              key: chartKeys[i],
              value: formatId(chartOffset, i, "rId"),
            });
          }
          const saPrefixes = ["smartart:", "smartart-lo:", "smartart-qs:", "smartart-cs:"];
          for (let i = 0; i < smartArtKeys.length; i++) {
            for (let p = 0; p < saPrefixes.length; p++) {
              entries.push({
                prefix: saPrefixes[p],
                key: smartArtKeys[i],
                value: formatId(smartArtOffset + p * smartArtKeys.length, i, "rId"),
              });
            }
          }
          for (const { reference, instance, numId } of ctx.numbering.concreteNumbering) {
            entries.push({ key: `${reference}-${instance}`, value: numId.toString() });
          }
          xmlData = replaceAllPlaceholders(xmlData, entries);
        } else {
          xmlData = replaceNumberingPlaceholders(xmlData, ctx.numbering.concreteNumbering);
        }
        return xmlData;
      })(),
      path: "word/document.xml",
    },
    Endnotes: {
      data: (() => {
        const endnoteCtx = mkCtx({
          relationships: ctx.endnotes.relationships,
        });
        const xmlData =
          XML_DECL +
          (endnotesDesc.stringify(
            {
              notes: ctx.endnotes.notes,
              separator: ctx.endnotes.separator,
              continuationSeparator: ctx.endnotes.continuationSeparator,
            },
            endnoteCtx,
          ) ?? "");
        const endnoteRelCount = ctx.endnotes.relationships.relationshipCount + 1;
        const endnoteMedia = findAndReplaceImagePlaceholders(
          xmlData,
          ctx.media.array,
          endnoteRelCount,
        );
        if (endnoteMedia.referenced.length > 0) {
          for (let i = 0; i < endnoteMedia.referenced.length; i++) {
            ctx.endnotes.relationships.addRelationship(
              endnoteRelCount + i,
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
              `media/${endnoteMedia.referenced[i].fileName}`,
            );
          }
          return replaceNumberingPlaceholders(endnoteMedia.xml, ctx.numbering.concreteNumbering);
        }
        return replaceNumberingPlaceholders(xmlData, ctx.numbering.concreteNumbering);
      })(),
      path: "word/endnotes.xml",
    },
    EndnotesRelationships:
      ctx.endnotes.relationships.relationshipCount > 0
        ? {
            data: XML_DECL + ctx.endnotes.relationships.serialize(),
            path: "word/_rels/endnotes.xml.rels",
          }
        : undefined,
    FileRelationships: {
      data: XML_DECL + ctx.fileRelationships.serialize(),
      path: "_rels/.rels",
    },
    FontTable: {
      data:
        XML_DECL +
        (fontTableDesc.stringify({ fonts: ctx.fontTable.fontOptionsWithKey }, ctx) ?? ""),
      path: "word/fontTable.xml",
    },
    FontTableRelationships: {
      data: XML_DECL + ctx.fontTable.relationships.serialize(),
      path: "word/_rels/fontTable.xml.rels",
    },
    FootNotes: {
      data: (() => {
        const xmlData = footnoteMedia.referenced.length > 0 ? footnoteMedia.xml : footnoteXmlData;
        return replaceNumberingPlaceholders(xmlData, ctx.numbering.concreteNumbering);
      })(),
      path: "word/footnotes.xml",
    },
    FootNotesRelationships:
      ctx.footNotes.relationships.relationshipCount > 0
        ? {
            data: XML_DECL + ctx.footNotes.relationships.serialize(),
            path: "word/_rels/footnotes.xml.rels",
          }
        : undefined,
    FooterRelationships: ctx.footers.map((entry, index) => {
      const footerCtx = mkCtx({ relationships: entry.relationships });
      const xmlData =
        XML_DECL + stringifyHeaderFooter("w:ftr", FOOTER_NAMESPACES, entry.children, footerCtx);
      footerFormattedViews.set(index, xmlData);
      // Footer images get per-part relationship IDs starting at
      // relationshipCount+1, mirroring the document part. The placeholder pass
      // uses referenced-local positions, so body r:embed and .rels stay aligned.
      const footerRelCount = entry.relationships.relationshipCount + 1;
      const footerMedia = findAndReplaceImagePlaceholders(xmlData, ctx.media.array, footerRelCount);
      footerMediaResults.set(index, footerMedia);

      for (let i = 0; i < footerMedia.referenced.length; i++) {
        entry.relationships.addRelationship(
          footerRelCount + i,
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
          `media/${footerMedia.referenced[i].fileName}`,
        );
      }

      return {
        data: XML_DECL + entry.relationships.serialize(),
        path: `word/_rels/footer${index + 1}.xml.rels`,
      };
    }),
    Footers: ctx.footers.map((_entry, index) => {
      const footerMedia = footerMediaResults.get(index)!;
      const tempXmlData = footerFormattedViews.get(index)!;
      const xmlData = footerMedia.referenced.length > 0 ? footerMedia.xml : tempXmlData;

      return {
        data: replaceNumberingPlaceholders(xmlData, ctx.numbering.concreteNumbering),
        path: `word/footer${index + 1}.xml`,
      };
    }),
    HeaderRelationships: ctx.headers.map((entry, index) => {
      const headerCtx = mkCtx({ relationships: entry.relationships });
      const xmlData =
        XML_DECL + stringifyHeaderFooter("w:hdr", HEADER_NAMESPACES, entry.children, headerCtx);
      headerFormattedViews.set(index, xmlData);
      // Header images get per-part relationship IDs starting at
      // relationshipCount+1, mirroring the document part. The placeholder pass
      // uses referenced-local positions, so body r:embed and .rels stay aligned.
      const headerRelCount = entry.relationships.relationshipCount + 1;
      const headerMedia = findAndReplaceImagePlaceholders(xmlData, ctx.media.array, headerRelCount);
      headerMediaResults.set(index, headerMedia);

      for (let i = 0; i < headerMedia.referenced.length; i++) {
        entry.relationships.addRelationship(
          headerRelCount + i,
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
          `media/${headerMedia.referenced[i].fileName}`,
        );
      }

      return {
        data: XML_DECL + entry.relationships.serialize(),
        path: `word/_rels/header${index + 1}.xml.rels`,
      };
    }),
    Headers: ctx.headers.map((_entry, index) => {
      const headerMedia = headerMediaResults.get(index)!;
      const tempXmlData = headerFormattedViews.get(index)!;
      const xmlData = headerMedia.referenced.length > 0 ? headerMedia.xml : tempXmlData;

      return {
        data: replaceNumberingPlaceholders(xmlData, ctx.numbering.concreteNumbering),
        path: `word/header${index + 1}.xml`,
      };
    }),
    Numbering: {
      data: ctx.numbering.serialize(),
      path: "word/numbering.xml",
    },
    Properties: {
      data: XML_DECL + (corePropertiesDesc.stringify(ctx._options, ctx) ?? ""),
      path: "docProps/core.xml",
    },
    Relationships: {
      data: (() => {
        for (let i = 0; i < documentMedia.referenced.length; i++) {
          ctx.document.relationships.addRelationship(
            documentRelationshipCount + i,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            `media/${documentMedia.referenced[i].fileName}`,
          );
        }

        const chartOffset = documentRelationshipCount + documentMedia.referenced.length;
        for (let i = 0; i < ctx.charts.array.length; i++) {
          ctx.document.relationships.addRelationship(
            chartOffset + i,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
            `charts/chart${i + 1}.xml`,
          );
        }

        addSmartArtRelationships(
          ctx.smartArts.array.map((s) => s.key),
          (id, type, target) => {
            ctx.document.relationships.addRelationship(id, type, target);
          },
          documentRelationshipCount + documentMedia.referenced.length + ctx.charts.array.length,
          0,
          {
            pathPrefix: "",
            styleRelType: "http://schemas.microsoft.com/office/2007/relationships/diagramStyle",
          },
        );

        ctx.document.relationships.addRelationship(
          ctx.document.relationships.relationshipCount + 1,
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
          "fontTable.xml",
        );

        return XML_DECL + ctx.document.relationships.serialize();
      })(),
      path: "word/_rels/document.xml.rels",
    },
    Settings: {
      data: XML_DECL + (settingsDesc.stringify(ctx._settingsOptions, ctx) ?? ""),
      path: "word/settings.xml",
    },
    Styles: {
      data: (() => {
        const xmlStyles = ctx.styles.serialize();
        return replaceNumberingPlaceholders(xmlStyles, ctx.numbering.concreteNumbering);
      })(),
      path: "word/styles.xml",
    },
    ...(ctx._options.bibliography
      ? {
          Bibliography: {
            data: XML_DECL + (bibliographyDesc.stringify(ctx._options.bibliography, ctx) ?? ""),
            path: "word/bibliography.xml",
          },
        }
      : {}),
    ...(ctx.charts.array.length > 0
      ? {
          Charts: ctx.charts.array.flatMap((chartData, i) => [
            {
              data: XML_DECL + chartData.chartSpaceXml,
              path: `word/charts/chart${i + 1}.xml`,
            },
            {
              data: '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
              path: `word/charts/_rels/chart${i + 1}.xml.rels`,
            },
          ]),
        }
      : {}),
    ...(ctx.smartArts.array.length > 0
      ? {
          DiagramData: ctx.smartArts.array.map((smartArtData, i) => ({
            data: XML_DECL + smartArtData.dataModelXml,
            path: `word/diagrams/data${i + 1}.xml`,
          })),
          DiagramLayout: ctx.smartArts.array.map((smartArtData, i) => ({
            data: getLayoutXml(smartArtData.layout),
            path: `word/diagrams/layout${i + 1}.xml`,
          })),
          DiagramStyle: ctx.smartArts.array.map((smartArtData, i) => ({
            data: getStyleXml(smartArtData.style),
            path: `word/diagrams/quickStyle${i + 1}.xml`,
          })),
          DiagramColors: ctx.smartArts.array.map((smartArtData, i) => ({
            data: getColorXml(smartArtData.color),
            path: `word/diagrams/colors${i + 1}.xml`,
          })),
          DiagramDrawing: ctx.smartArts.array.map((_, i) => ({
            data: DEFAULT_DRAWING_XML,
            path: `word/diagrams/drawing${i + 1}.xml`,
          })),
        }
      : {}),
    ...(ctx.altChunks.array.length > 0
      ? {
          AltChunks: ctx.altChunks.array.map((altChunkData) => ({
            data: altChunkData.data,
            path: `word/${altChunkData.path}`,
          })),
        }
      : {}),
    ...(ctx.subDocs.array.length > 0
      ? {
          SubDocs: ctx.subDocs.array.map((subDocData) => ({
            data: subDocData.data,
            path: `word/${subDocData.path}`,
          })),
        }
      : {}),
    ...(ctx.glossaryOptions
      ? {
          Glossary: {
            data: (() => {
              const glossaryCtx = mkCtx(undefined);
              return XML_DECL + (glossaryDesc.stringify(ctx.glossaryOptions!, glossaryCtx) ?? "");
            })(),
            path: "word/glossary/document.xml",
          },
        }
      : {}),
    ...(ctx.webSettings
      ? {
          WebSettings: {
            data: XML_DECL + (webSettingsDesc.stringify(ctx._options.webSettings ?? {}, ctx) ?? ""),
            path: "word/webSettings.xml",
          },
        }
      : {}),
  };
}
