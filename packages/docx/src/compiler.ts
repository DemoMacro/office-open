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
  addBinaryFile,
  addSmartArtRelationships,
  compileMapping,
  contentTypesDesc,
  createThemeXml,
  deriveContentTypes,
  DOCX_PARTS,
  findAndReplaceImagePlaceholders,
  formatId,
  hasPlaceholders,
  optionalRelsPart,
  Relationships,
  TargetModeType,
  replaceAllPlaceholders,
  replaceNumberingPlaceholders,
  IMAGE_MEDIA_CONTENT_TYPES,
  resolverFromRegistry,
  toUint8Array,
} from "@office-open/core";
import type { XmlifyedFile, Zippable } from "@office-open/core";
import {
  DEFAULT_DRAWING_XML,
  getColorXml,
  getLayoutXml,
  getStyleXml,
  stringifyColorDefinitionPart,
  stringifyLayoutDefinitionPart,
  stringifyStyleDefinitionPart,
} from "@office-open/core/smartart";
import { OOXML_XML_DECLARATION, escapeXml } from "@office-open/xml";
import type { DocumentOptions } from "@parts/core-properties";
import { obfuscate } from "@parts/fonts/obfuscate-ttf-to-odttf";
import { HEADER_NAMESPACES, FOOTER_NAMESPACES, stringifyHeaderFooter } from "@parts/header-footer";
import type { CommentOptions } from "@parts/paragraph/run/comment-run";

import { stringifyDocumentXml, stringifyBodyChild, type BodyContext } from "./body";
import { DocxWriteContext } from "./context";
import {
  corePropertiesDesc,
  customPropertiesDesc,
  appPropertiesDesc,
  fontTableDesc,
  webSettingsDesc,
  commentsDesc,
  commentsExtendedDesc,
  peopleDesc,
  bibliographyDesc,
  settingsDesc,
  footnotesDesc,
  endnotesDesc,
  glossaryDesc,
} from "./parts";

/** Reusable TextEncoder (stateless, safe to share). */
const encoder = new TextEncoder();

/** Relationship type for OLE embedding parts (word|ppt/embeddings/*). */
const OLE_OBJECT_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";

/** XML declaration prepended to every OOXML part. */
const XML_DECL = OOXML_XML_DECLARATION;

/** DOCX part path → content type, derived from the part registry. */
const DOCX_CONTENT_TYPE_RESOLVER = resolverFromRegistry(DOCX_PARTS);

/** Extension → MIME for media/font/embedding Default entries. Declared only
 * for extensions actually present in the package. */
const DOCX_MEDIA_CONTENT_TYPES: Record<string, string> = {
  ...IMAGE_MEDIA_CONTENT_TYPES,
  odttf: "application/vnd.openxmlformats-officedocument.obfuscatedFont",
  bin: "application/vnd.openxmlformats-officedocument.oleObject",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  xls: "application/vnd.ms-excel",
  xlsb: "application/vnd.ms-excel.sheet.binary.macroEnabled.12",
};

/** Extended context for header/footer part stringification. */
type DocxContext = BodyContext;

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
  const xmlifiedFileMapping = xmlifyContext(ctx);
  const files = compileMapping(xmlifiedFileMapping, overrides);

  // Media files
  const mediaArray = ctx.media.array;
  for (const mediaData of mediaArray) {
    addBinaryFile(files, `word/media/${mediaData.fileName}`, mediaData.data, mediaLevel);
    if (mediaData.type === "svg") {
      addBinaryFile(
        files,
        `word/media/${mediaData.fallback.fileName}`,
        mediaData.fallback.data,
        mediaLevel,
      );
    }
  }

  // OLE embedding binaries (word/embeddings/oleObjectN.bin)
  for (const embedding of ctx.embeddings.array) {
    addBinaryFile(files, `word/embeddings/${embedding.fileName}`, embedding.data, mediaLevel);
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
    files[part.path] = toUint8Array(part.data);
  }

  // [Content_Types].xml is serialized last: parts register their media/fonts
  // during stringify (run by xmlifyContext above), so backfilling <Default>
  // extensions from `ctx` now sees the complete set. Building it inside
  // xmlifyContext's object literal evaluated it before header/footer/font media
  // was registered, leaving jpg/gif/odttf without a covering Default.
  files["[Content_Types].xml"] = encoder.encode(buildContentTypesData(ctx, files));

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
  Numbering?: XmlifyedFile;
  NumberingRelationships?: XmlifyedFile;
  Relationships: XmlifyedFile;
  FileRelationships: XmlifyedFile;
  Headers: XmlifyedFile[];
  Footers: XmlifyedFile[];
  HeaderRelationships: XmlifyedFile[];
  FooterRelationships: XmlifyedFile[];
  CustomProperties?: XmlifyedFile;
  AppProperties: XmlifyedFile;
  FootNotes?: XmlifyedFile;
  FootNotesRelationships?: XmlifyedFile;
  Endnotes?: XmlifyedFile;
  EndnotesRelationships?: XmlifyedFile;
  Settings: XmlifyedFile;
  Comments?: XmlifyedFile;
  CommentsRelationships?: XmlifyedFile;
  People?: XmlifyedFile;
  CommentsExtended?: XmlifyedFile;
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

/**
 * Comments carried by the document: those the caller listed explicitly
 * (`options.comments`) plus entries registered by `{ comment }` sugar children
 * during body stringification. Drives both word/comments.xml generation and the
 * [Content_Types] comments Override, which must stay in sync (OPC consistency).
 */
function mergedCommentChildren(ctx: DocxWriteContext): CommentOptions[] {
  return [...(ctx._options.comments ?? []), ...ctx.comments.entries];
}

/**
 * Serialize [Content_Types].xml from the part registry, then backfill media/
 * font/embedding `<Default>` entries from the parts actually written.
 *
 * Must run after every part has been stringified (parts call `ctx.addMedia`
 * during stringify), so call this once `xmlifyContext` has finished — not from
 * inside its object literal, where ContentTypes would evaluate before the
 * later-defined header/footer/font parts have registered their media.
 */
/**
 * Derive [Content_Types].xml from the parts actually written to the package.
 *
 * The file set is the single source of truth: every part the resolver knows
 * becomes an Override, every other file falls through to an extension Default,
 * and altChunk/sub-document parts — whose content type is data-driven, not
 * path-determinable — are injected as explicit overrides.
 *
 * Round-trip no longer passes the source [Content_Types] through: the compiler
 * regenerates every part path (altChunks get a fresh uniqueId), so deriving
 * from the written files is what keeps declarations and parts in sync.
 */
function buildContentTypesData(ctx: DocxWriteContext, files: Zippable): string {
  const overrides = [
    ...ctx.altChunks.array.map((ac) => ({
      path: `word/${ac.path}`,
      contentType: ac.contentType ?? "application/xhtml+xml",
    })),
    ...ctx.subDocs.array.map((sd) => ({
      path: `word/${sd.path}`,
      contentType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    })),
  ];
  const input = deriveContentTypes(Object.keys(files), {
    resolve: DOCX_CONTENT_TYPE_RESOLVER,
    mediaContentTypes: DOCX_MEDIA_CONTENT_TYPES,
    overrides,
  });
  return XML_DECL + (contentTypesDesc.stringify(input, ctx) ?? "");
}

function xmlifyContext(ctx: DocxWriteContext): XmlifyedFileMapping {
  const mkCtx = (viewWrapper: DocxContext["viewWrapper"] = ctx.document): DocxContext => {
    const bodyCtx: DocxContext = {
      fileData: ctx,
      file: ctx,
      viewWrapper,
      addRelationship: (type: string, target: string, mode?: string) =>
        ctx.addRelationship(type, target, mode),
      addMedia: (data: Uint8Array, type: string) => ctx.addMedia(data, type),
      addHyperlink: (key, target) => ctx.addHyperlink(key, target),
      // Assigned after the literal: stringifyBodyChild needs this context, which
      // is only bound once the literal finishes initializing.
      stringifyChild: undefined as unknown as DocxContext["stringifyChild"],
    };
    bodyCtx.stringifyChild = (child) => stringifyBodyChild(child, bodyCtx);
    return bodyCtx;
  };

  const documentRelationshipCount = ctx.document.relationships.relationshipCount + 1;
  // Per-part media-replacement results shared between the .rels pass and the
  // body-XML pass so both use identical rId offsets. Each header/footer part
  // has its own relationship numbering (independent of the document part).
  // Embedding results chain after media (same {fileName} placeholder bridge,
  // offsets continuing past the media relationships), mirroring the document
  // part — headers/footers can carry w:object runs of their own.
  const footerMediaResults = new Map<number, { xml: string; referenced: { fileName: string }[] }>();
  const footerEmbeddingResults = new Map<
    number,
    { xml: string; referenced: { fileName: string }[] }
  >();
  const headerMediaResults = new Map<number, { xml: string; referenced: { fileName: string }[] }>();
  const headerEmbeddingResults = new Map<
    number,
    { xml: string; referenced: { fileName: string }[] }
  >();
  const docCtx = mkCtx(ctx.document);
  const documentXmlData = XML_DECL + stringifyDocumentXml(ctx, docCtx);

  // Comments is an optional part — skip it entirely (no comments.xml, no
  // comments rels, no [Content_Types] Override) when the document carries none.
  // Emitting an empty comments.xml with a dangling relationship is the OPC
  // violation that makes Word reject the package on open.
  const mergedCommentChildrenList = mergedCommentChildren(ctx);
  const hasComments = mergedCommentChildrenList.length > 0;
  const commentRelationshipCount = hasComments
    ? ctx.comments.relationships.relationshipCount + 1
    : 0;
  const commentCtx = hasComments ? mkCtx({ relationships: ctx.comments.relationships }) : null;
  const commentXmlData = commentCtx
    ? XML_DECL + commentsDesc.stringify(mergedCommentChildrenList, commentCtx)
    : "";

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
  // Sampled after stringify: hyperlinks registered during note stringification
  // take sequential ids here, and the media/embedding offsets below must skip
  // them (same ordering as the document and comments parts).
  const footnoteRelationshipCount = ctx.footNotes.relationships.relationshipCount + 1;

  const documentMedia = findAndReplaceImagePlaceholders(
    documentXmlData,
    ctx.media.array,
    documentRelationshipCount,
  );
  // OLE embeddings reuse the same {fileName} placeholder bridge as images; run
  // after media so {oleObjectN.bin} placeholders resolve against the embedding array.
  const documentEmbeddingOffset = documentRelationshipCount + documentMedia.referenced.length;
  const documentEmbeddings = findAndReplaceImagePlaceholders(
    documentMedia.xml,
    ctx.embeddings.array,
    documentEmbeddingOffset,
  );
  const commentMedia = hasComments
    ? findAndReplaceImagePlaceholders(commentXmlData, ctx.media.array, commentRelationshipCount)
    : { xml: "", referenced: [] as { fileName: string }[] };
  // OLE embeddings inside comments reuse the same placeholder bridge, offset
  // chained past the media relationships (same order as the document part).
  const commentEmbeddingOffset = commentRelationshipCount + commentMedia.referenced.length;
  const commentEmbeddings = hasComments
    ? findAndReplaceImagePlaceholders(
        commentMedia.xml,
        ctx.embeddings.array,
        commentEmbeddingOffset,
      )
    : { xml: "", referenced: [] as { fileName: string }[] };
  const footnoteMedia = findAndReplaceImagePlaceholders(
    footnoteXmlData,
    ctx.media.array,
    footnoteRelationshipCount,
  );
  // OLE embeddings inside footnotes reuse the same placeholder bridge, offset
  // chained past the media relationships (same order as the document part).
  const footnoteEmbeddingOffset = footnoteRelationshipCount + footnoteMedia.referenced.length;
  const footnoteEmbeddings = findAndReplaceImagePlaceholders(
    footnoteMedia.xml,
    ctx.embeddings.array,
    footnoteEmbeddingOffset,
  );
  // Register footnote media/embedding relationships eagerly so the
  // relationshipCount used to gate footnotes.xml.rels reflects the final state
  // (see FootNotesRelationships).
  for (const [i, ref] of footnoteMedia.referenced.entries()) {
    ctx.footNotes.relationships.addRelationship(
      footnoteRelationshipCount + i,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
      `media/${ref.fileName}`,
    );
  }
  for (const [i, ref] of footnoteEmbeddings.referenced.entries()) {
    ctx.footNotes.relationships.addRelationship(
      footnoteEmbeddingOffset + i,
      OLE_OBJECT_RELATIONSHIP,
      `embeddings/${ref.fileName}`,
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
            data: replaceNumberingPlaceholders(
              commentEmbeddings.xml,
              ctx.numbering.concreteNumbering,
            ),
            path: "word/comments.xml",
          },
          CommentsRelationships: (() => {
            for (const [i, ref] of commentMedia.referenced.entries()) {
              ctx.comments.relationships.addRelationship(
                commentRelationshipCount + i,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                `media/${ref.fileName}`,
              );
            }
            for (const [i, ref] of commentEmbeddings.referenced.entries()) {
              ctx.comments.relationships.addRelationship(
                commentEmbeddingOffset + i,
                OLE_OBJECT_RELATIONSHIP,
                `embeddings/${ref.fileName}`,
              );
            }
            return optionalRelsPart(
              ctx.comments.relationships,
              XML_DECL,
              "word/_rels/comments.xml.rels",
            );
          })(),
        }
      : {}),
    // docProps/custom.xml — emitted only when custom properties exist (parsed
    // presence or fresh authoring); Word omits the part otherwise.
    ...(ctx._options.customProperties !== undefined
      ? {
          CustomProperties: {
            data:
              XML_DECL +
              (customPropertiesDesc.stringify({ properties: ctx._options.customProperties }, ctx) ??
                ""),
            path: "docProps/custom.xml",
          },
        }
      : {}),
    // Word 2013+ comment infrastructure — emitted only when populated, same
    // conditional rule as comments.xml (an empty part with a relationship is
    // an OPC violation Word rejects).
    ...(ctx._options.people?.length
      ? {
          People: {
            data: XML_DECL + (peopleDesc.stringify(ctx._options.people, ctx) ?? ""),
            path: "word/people.xml",
          },
        }
      : {}),
    ...(ctx._options.commentsExtended?.length
      ? {
          CommentsExtended: {
            data:
              XML_DECL + (commentsExtendedDesc.stringify(ctx._options.commentsExtended, ctx) ?? ""),
            path: "word/commentsExtended.xml",
          },
        }
      : {}),
    Document: {
      data: (() => {
        let xmlData = documentEmbeddings.xml;
        if (hasPlaceholders(xmlData)) {
          const mediaCount = documentMedia.referenced.length;
          const embeddingCount = documentEmbeddings.referenced.length;
          const chartKeys = ctx.charts.array.map((c) => c.key);
          const smartArtKeys = ctx.smartArts.array.map((s) => s.key);
          const chartOffset = documentRelationshipCount + mediaCount + embeddingCount;
          const smartArtOffset = chartOffset + chartKeys.length;

          // Build combined replacement entries for charts, smartart, and numbering
          const entries: Array<{ prefix?: string; key: string; value: string }> = [];
          for (const [i, key] of chartKeys.entries()) {
            entries.push({
              prefix: "chart:",
              key,
              value: formatId(chartOffset, i, "rId"),
            });
          }
          const saPrefixes = ["smartart:", "smartart-lo:", "smartart-qs:", "smartart-cs:"];
          for (const [i, key] of smartArtKeys.entries()) {
            for (let p = 0; p < saPrefixes.length; p++) {
              entries.push({
                prefix: saPrefixes[p],
                key,
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
    // Theme — fresh-compile emits a language-neutral default theme
    // (createThemeXml). Round-trip carries the source theme in rawParts,
    // already copied verbatim above, so skip emitting here to avoid a duplicate.
    ...(ctx._options.rawParts?.some((part) => part.path.startsWith("word/theme/"))
      ? {}
      : {
          Theme: {
            data: XML_DECL + createThemeXml(),
            path: "word/theme/theme1.xml",
          },
        }),
    ...(ctx.hasEndnotes
      ? {
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
              for (const [i, ref] of endnoteMedia.referenced.entries()) {
                ctx.endnotes.relationships.addRelationship(
                  endnoteRelCount + i,
                  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                  `media/${ref.fileName}`,
                );
              }
              // OLE embeddings inside endnotes reuse the same placeholder
              // bridge, offset chained past the media relationships.
              const endnoteEmbeddingOffset = endnoteRelCount + endnoteMedia.referenced.length;
              const endnoteEmbeddings = findAndReplaceImagePlaceholders(
                endnoteMedia.xml,
                ctx.embeddings.array,
                endnoteEmbeddingOffset,
              );
              for (const [i, ref] of endnoteEmbeddings.referenced.entries()) {
                ctx.endnotes.relationships.addRelationship(
                  endnoteEmbeddingOffset + i,
                  OLE_OBJECT_RELATIONSHIP,
                  `embeddings/${ref.fileName}`,
                );
              }
              return replaceNumberingPlaceholders(
                endnoteEmbeddings.xml,
                ctx.numbering.concreteNumbering,
              );
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
        }
      : {}),
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
    FontTableRelationships: optionalRelsPart(
      ctx.fontTable.relationships,
      XML_DECL,
      "word/_rels/fontTable.xml.rels",
    ),
    ...(ctx.hasFootnotes
      ? {
          FootNotes: {
            data: replaceNumberingPlaceholders(
              footnoteEmbeddings.xml,
              ctx.numbering.concreteNumbering,
            ),
            path: "word/footnotes.xml",
          },
          FootNotesRelationships:
            ctx.footNotes.relationships.relationshipCount > 0
              ? {
                  data: XML_DECL + ctx.footNotes.relationships.serialize(),
                  path: "word/_rels/footnotes.xml.rels",
                }
              : undefined,
        }
      : {}),
    FooterRelationships: ctx.footers
      .map((entry, index) => {
        const footerCtx = mkCtx({ relationships: entry.relationships });
        const xmlData =
          XML_DECL + stringifyHeaderFooter("w:ftr", FOOTER_NAMESPACES, entry.children, footerCtx);
        // Footer images get per-part relationship IDs starting at
        // relationshipCount+1, mirroring the document part. The placeholder pass
        // uses referenced-local positions, so body r:embed and .rels stay aligned.
        const footerRelCount = entry.relationships.relationshipCount + 1;
        const footerMedia = findAndReplaceImagePlaceholders(
          xmlData,
          ctx.media.array,
          footerRelCount,
        );
        footerMediaResults.set(index, footerMedia);
        // OLE embeddings reuse the same {fileName} placeholder bridge, offset
        // chained past the media relationships (same order as the document part).
        const footerEmbeddingOffset = footerRelCount + footerMedia.referenced.length;
        const footerEmbeddings = findAndReplaceImagePlaceholders(
          footerMedia.xml,
          ctx.embeddings.array,
          footerEmbeddingOffset,
        );
        footerEmbeddingResults.set(index, footerEmbeddings);

        for (const [i, ref] of footerMedia.referenced.entries()) {
          entry.relationships.addRelationship(
            footerRelCount + i,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            `media/${ref.fileName}`,
          );
        }
        for (const [i, ref] of footerEmbeddings.referenced.entries()) {
          entry.relationships.addRelationship(
            footerEmbeddingOffset + i,
            OLE_OBJECT_RELATIONSHIP,
            `embeddings/${ref.fileName}`,
          );
        }

        return optionalRelsPart(
          entry.relationships,
          XML_DECL,
          `word/_rels/${entry.partName ?? `footer${index + 1}.xml`}.rels`,
        );
      })
      .filter((r): r is XmlifyedFile => r !== undefined),
    Footers: ctx.footers.map((entry, index) => {
      const footerEmbeddings = footerEmbeddingResults.get(index)!;

      return {
        data: replaceNumberingPlaceholders(footerEmbeddings.xml, ctx.numbering.concreteNumbering),
        path: `word/${entry.partName ?? `footer${index + 1}.xml`}`,
      };
    }),
    HeaderRelationships: ctx.headers
      .map((entry, index) => {
        const headerCtx = mkCtx({ relationships: entry.relationships });
        const xmlData =
          XML_DECL + stringifyHeaderFooter("w:hdr", HEADER_NAMESPACES, entry.children, headerCtx);
        // Header images get per-part relationship IDs starting at
        // relationshipCount+1, mirroring the document part. The placeholder pass
        // uses referenced-local positions, so body r:embed and .rels stay aligned.
        const headerRelCount = entry.relationships.relationshipCount + 1;
        const headerMedia = findAndReplaceImagePlaceholders(
          xmlData,
          ctx.media.array,
          headerRelCount,
        );
        headerMediaResults.set(index, headerMedia);
        // OLE embeddings reuse the same {fileName} placeholder bridge, offset
        // chained past the media relationships (same order as the document part).
        const headerEmbeddingOffset = headerRelCount + headerMedia.referenced.length;
        const headerEmbeddings = findAndReplaceImagePlaceholders(
          headerMedia.xml,
          ctx.embeddings.array,
          headerEmbeddingOffset,
        );
        headerEmbeddingResults.set(index, headerEmbeddings);

        for (const [i, ref] of headerMedia.referenced.entries()) {
          entry.relationships.addRelationship(
            headerRelCount + i,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            `media/${ref.fileName}`,
          );
        }
        for (const [i, ref] of headerEmbeddings.referenced.entries()) {
          entry.relationships.addRelationship(
            headerEmbeddingOffset + i,
            OLE_OBJECT_RELATIONSHIP,
            `embeddings/${ref.fileName}`,
          );
        }

        return optionalRelsPart(
          entry.relationships,
          XML_DECL,
          `word/_rels/${entry.partName ?? `header${index + 1}.xml`}.rels`,
        );
      })
      .filter((r): r is XmlifyedFile => r !== undefined),
    Headers: ctx.headers.map((entry, index) => {
      const headerEmbeddings = headerEmbeddingResults.get(index)!;

      return {
        data: replaceNumberingPlaceholders(headerEmbeddings.xml, ctx.numbering.concreteNumbering),
        path: `word/${entry.partName ?? `header${index + 1}.xml`}`,
      };
    }),
    ...(ctx.hasNumbering
      ? (() => {
          // Picture-bullet media resolves through numbering.xml's own rels —
          // same placeholder bridge as headers/footers, ids local to the part
          // (numbering.xml carries no other relationships).
          const numberingXml = ctx.numbering.serialize(ctx);
          const numberingMedia = findAndReplaceImagePlaceholders(numberingXml, ctx.media.array, 1);
          const numberingRels = new Relationships();
          for (const [i, ref] of numberingMedia.referenced.entries()) {
            numberingRels.addRelationship(
              1 + i,
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
              `media/${ref.fileName}`,
            );
          }
          return {
            Numbering: { data: numberingMedia.xml, path: "word/numbering.xml" },
            NumberingRelationships: optionalRelsPart(
              numberingRels,
              XML_DECL,
              "word/_rels/numbering.xml.rels",
            ),
          };
        })()
      : {}),
    Properties: {
      data: XML_DECL + (corePropertiesDesc.stringify(ctx._options, ctx) ?? ""),
      path: "docProps/core.xml",
    },
    Relationships: {
      data: (() => {
        for (const [i, ref] of documentMedia.referenced.entries()) {
          ctx.document.relationships.addRelationship(
            documentRelationshipCount + i,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            `media/${ref.fileName}`,
          );
        }
        for (const [i, ref] of documentEmbeddings.referenced.entries()) {
          ctx.document.relationships.addRelationship(
            documentEmbeddingOffset + i,
            OLE_OBJECT_RELATIONSHIP,
            `embeddings/${ref.fileName}`,
          );
        }

        const chartOffset =
          documentRelationshipCount +
          documentMedia.referenced.length +
          documentEmbeddings.referenced.length;
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
          documentRelationshipCount +
            documentMedia.referenced.length +
            documentEmbeddings.referenced.length +
            ctx.charts.array.length,
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
    ...(ctx._settingsOptions.attachedTemplate !== undefined
      ? (() => {
          const rels = new Relationships();
          rels.addRelationship(
            1,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate",
            ctx._settingsOptions.attachedTemplate,
            TargetModeType.EXTERNAL,
          );
          return {
            SettingsRelationships: {
              data: XML_DECL + rels.serialize(),
              path: "word/_rels/settings.xml.rels",
            },
          };
        })()
      : {}),
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
          Charts: ctx.charts.array.map((chartData, i) => ({
            data: XML_DECL + chartData.chartSpaceXml,
            path: `word/charts/chart${i + 1}.xml`,
          })),
          // Embedded workbooks behind c:externalData: each chart's own rels
          // part plus the word/embeddings binary. The relationship id is
          // carried verbatim so the re-emitted r:id resolves without rewriting.
          ...(ctx.charts.array.some((c) => c.embedding)
            ? {
                ChartRels: ctx.charts.array.flatMap((chartData, i) => {
                  const e = chartData.embedding;
                  if (!e) return [];
                  const relsXml =
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                    `<Relationship Id="${escapeXml(e.relationshipId)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/${escapeXml(e.fileName)}"/>` +
                    "</Relationships>";
                  return [
                    {
                      data: XML_DECL + relsXml,
                      path: `word/charts/_rels/chart${i + 1}.xml.rels`,
                    },
                  ];
                }),
                ChartEmbeddings: (() => {
                  const seen = new Set<string>();
                  const parts: { data: Uint8Array; path: string }[] = [];
                  for (const c of ctx.charts.array) {
                    const e = c.embedding;
                    if (!e || seen.has(e.fileName)) continue;
                    seen.add(e.fileName);
                    parts.push({
                      data: toUint8Array(e.data),
                      path: `word/embeddings/${e.fileName}`,
                    });
                  }
                  return parts;
                })(),
              }
            : {}),
        }
      : {}),
    ...(ctx.smartArts.array.length > 0
      ? {
          DiagramData: ctx.smartArts.array.map((smartArtData, i) => ({
            data: XML_DECL + smartArtData.dataModelXml,
            path: `word/diagrams/data${i + 1}.xml`,
          })),
          DiagramLayout: ctx.smartArts.array.map((smartArtData, i) => ({
            data:
              typeof smartArtData.layout === "string"
                ? getLayoutXml(smartArtData.layout)
                : stringifyLayoutDefinitionPart(smartArtData.layout),
            path: `word/diagrams/layout${i + 1}.xml`,
          })),
          DiagramStyle: ctx.smartArts.array.map((smartArtData, i) => ({
            data:
              typeof smartArtData.style === "string"
                ? getStyleXml(smartArtData.style)
                : stringifyStyleDefinitionPart(smartArtData.style),
            path: `word/diagrams/quickStyle${i + 1}.xml`,
          })),
          DiagramColors: ctx.smartArts.array.map((smartArtData, i) => ({
            data:
              typeof smartArtData.color === "string"
                ? getColorXml(smartArtData.color)
                : stringifyColorDefinitionPart(smartArtData.color),
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
