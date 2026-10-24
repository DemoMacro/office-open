/**
 * DOCX document compiler — pure function entry point.
 *
 * compileDocument() accepts DocumentOptions directly,
 * creates a DocxWriteContext internally, and produces a Zippable result.
 * All XML parts are produced via descriptors or serialize() —
 * no Formatter dependency.
 *
 * The part wiring runs in phases owned by the compile/ folder: the document
 * body and its rels (compile/document), the notes parts (compile/notes),
 * headers/footers (compile/headerfooter), and the chart/SmartArt tail
 * (compile/drawings), with the placeholder bridge and embedding rel
 * resolution they share in compile/shared. This module keeps the
 * part-mapping orchestration.
 *
 * @module
 */

import {
  addModelBinaries,
  compileMapping,
  createThemeXml,
  encodeUriPath,
  Relationships,
  replaceNumberingPlaceholders,
} from "@office-open/core";
import type { ReproducibleScope } from "@office-open/core";
import type { OoxmlPackageVariant } from "@office-open/core";
import type { WriteContext } from "@office-open/core/descriptor";
import {
  DOCX_PARTS,
  dropDanglingPassthroughRels,
  finalizeContentTypes,
  IMAGE_MEDIA_CONTENT_TYPES,
  ooxmlPackageFormatInfo,
  optionalRelsPart,
  partContentTypeResolver,
  partPathToRelsPath,
  RELATIONSHIP_TYPES,
  resolveMainPartContentType,
  TargetModeType,
} from "@office-open/core/opc";
import type { RelationshipType, XmlifyedFile, Zippable } from "@office-open/core/opc";
import { buildThemeXml } from "@office-open/core/theme";
import type { DocumentOptions } from "@parts/core-properties";
import {
  documentNamespaceDialect,
  type DocumentNamespaceDialect,
} from "@parts/document/document-attributes";
import { defaultEmbeddedFontPath, embeddedFontParts } from "@parts/fonts/font-wrapper";
import { obfuscate } from "@parts/fonts/obfuscate-ttf-to-odttf";
import type { MailMergeOptions } from "@parts/settings/settings";

import { stringifyDocumentXml, stringifyBodyChild, type BodyContext } from "./body";
import { compileDocumentEntries } from "./compile/document";
import { compileChartExParts, compileChartParts, compileSmartArtParts } from "./compile/drawings";
import { compileGlossaryParts } from "./compile/glossary";
import { compileHeaderFooterParts } from "./compile/headerfooter";
import { compileNotesParts } from "./compile/notes";
import { XML_DECL, resolveNumberingMedia } from "./compile/shared";
import { compileTextBoxParts } from "./compile/textbox";
import { DocxWriteContext, themePartName } from "./context";
import {
  corePropertiesDesc,
  customPropertiesDesc,
  customizationsDesc,
  appPropertiesDesc,
  fontTableDesc,
  webSettingsDesc,
  bibliographyDesc,
  settingsDesc,
  attachedTemplateRelationshipId,
  mailMergeRecipientsDesc,
  peopleDesc,
  commentsIdsDesc,
  commentsExtendedDesc,
  commentsExtensibleDesc,
} from "./parts";

/** Reusable TextEncoder (stateless, safe to share). */
const encoder = new TextEncoder();
const decoder = new TextDecoder();

/** DOCX part path → content type, derived from the part manifest. */
const DOCX_CONTENT_TYPE_RESOLVER = partContentTypeResolver(DOCX_PARTS);

const DOCX_THEME_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.theme+xml";
const DOCX_CHART_EX_CONTENT_TYPE = "application/vnd.ms-office.chartex+xml";

function bindThemeMedia(xml: string, ctx: DocxWriteContext, rels: Relationships): string {
  const names = new Set(ctx.media.array.map((media) => media.fileName));
  if (![...names].some((name) => xml.includes(`{${name}}`))) return xml;
  const ids = new Map<string, string>();
  return xml.replace(/\{([^{}]+)\}/g, (placeholder, fileName: string) => {
    const existing = ids.get(fileName);
    if (existing) return existing;
    if (!names.has(fileName)) return placeholder;
    const id = `rId${rels.add(RELATIONSHIP_TYPES.image, `../media/${fileName}`)}`;
    ids.set(fileName, id);
    return id;
  });
}

function withThemeRelationships(ctx: DocxWriteContext, rels: Relationships): WriteContext {
  const themeCtx = Object.create(ctx);
  themeCtx.addRelationship = (type: RelationshipType, target: string, mode?: string) => {
    const ownerPath = themePartName(ctx._options);
    const source = ownerPath.split("/").slice(0, -1);
    const targetPath = target.split("/").slice(0, -1);
    let common = 0;
    while (
      common < source.length &&
      common < targetPath.length &&
      source[common] === targetPath[common]
    )
      common++;
    const relativeTarget = [
      ...Array.from({ length: source.length - common }, () => ".."),
      ...target.split("/").slice(common),
    ].join("/");
    return `rId${rels.add(
      type,
      target.startsWith("/") ? target : relativeTarget,
      mode as "External" | undefined,
    )}`;
  };
  themeCtx.addMedia = ctx.addMedia.bind(ctx);
  return themeCtx;
}

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

function useStrictRelationshipTypes(files: Zippable, dialect: DocumentNamespaceDialect): void {
  if (dialect !== "strict") return;
  for (const [path, data] of Object.entries(files)) {
    if (!path.endsWith(".rels")) continue;
    const bytes =
      data instanceof Uint8Array
        ? data
        : Array.isArray(data) && data[0] instanceof Uint8Array
          ? data[0]
          : undefined;
    if (!bytes) continue;
    const xml = new TextDecoder()
      .decode(bytes)
      .replaceAll(
        /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\//g,
        'Type="http://purl.oclc.org/ooxml/officeDocument/relationships/',
      );
    files[path] = Array.isArray(data) ? [encoder.encode(xml), data[1]] : encoder.encode(xml);
  }
}

/** Extended context for header/footer part stringification. */
type DocxContext = BodyContext;

/** Factory the compile phases use to derive a per-part stringify context. */
export type PartCtxFactory = (viewWrapper?: BodyContext["viewWrapper"]) => BodyContext;

/** Source-backed part presence; absent source Content Types means fresh compile. */
function hasSourcePart(options: DocumentOptions, partName: string): boolean {
  const result =
    !options.contentTypes ||
    options.contentTypes.overrides.some(
      (override) => override.partName.toLowerCase() === `/${partName.toLowerCase()}`,
    );
  return result;
}

function sourcePartContentType(options: DocumentOptions, partName: string): string | undefined {
  return options.contentTypes?.overrides.find(
    (override) => override.partName.toLowerCase() === `/${partName.toLowerCase()}`,
  )?.contentType;
}

function canForcePartContentType(
  options: DocumentOptions,
  partName: string,
  contentType: string,
): boolean {
  const source = sourcePartContentType(options, partName);
  return source === undefined || source === contentType;
}

function renamePrimaryDocument(files: Zippable, sourcePath: string, targetPath: string): void {
  if (sourcePath === targetPath) return;
  const sourceRelsPath = partPathToRelsPath(sourcePath);
  const targetRelsPath = partPathToRelsPath(targetPath);
  if (files[sourcePath] !== undefined) {
    files[targetPath] = files[sourcePath];
    delete files[sourcePath];
  }
  if (files[sourceRelsPath] !== undefined) {
    files[targetRelsPath] = files[sourceRelsPath];
    delete files[sourceRelsPath];
  }
  for (const path of ["_rels/.rels", "[Content_Types].xml"]) {
    const data = files[path];
    if (!(data instanceof Uint8Array)) continue;
    let xml = decoder.decode(data);
    xml = xml
      .replaceAll(`Target="/${sourcePath}"`, `Target="${targetPath}"`)
      .replaceAll(`Target="${sourcePath}"`, `Target="${targetPath}"`)
      .replaceAll(`PartName="/${sourcePath}"`, `PartName="/${targetPath}"`);
    files[path] = encoder.encode(xml);
  }
}

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
  reproducible?: ReproducibleScope,
  packageVariant: OoxmlPackageVariant = "standard",
): Zippable {
  const packageFormat = ooxmlPackageFormatInfo("wordprocessing", packageVariant);
  const mainPartPath = options.primaryPartPath ?? packageFormat.mainPartPath;
  const ctx = new DocxWriteContext(options, reproducible);
  if (options.mailMergeRecipients) {
    const settings = ctx._settingsOptions;
    const odso = settings.mailMerge?.odso;
    const recipientData = [...(odso?.recipientData ?? [])];
    const sourceRelationships = options.passthroughRelationships ?? [];
    const attachedTemplateId =
      settings.attachedTemplate === undefined
        ? undefined
        : attachedTemplateRelationshipId(
            settings.attachedTemplate,
            sourceRelationships,
            recipientData,
          );
    const occupiedIds = new Set([
      ...sourceRelationships
        .filter((relationship) => relationship.source === "word/settings.xml")
        .map((relationship) => relationship.rId),
      ...(attachedTemplateId === undefined ? [] : [attachedTemplateId]),
    ]);
    const mailMerge = (settings.mailMerge ?? {
      mainDocumentType: "letter",
      dataType: "database",
    }) as MailMergeOptions;
    for (let i = recipientData.length; i < options.mailMergeRecipients.length; i++) {
      let id = 1;
      while (occupiedIds.has(`rId${id}`) || recipientData.includes(`rId${id}`)) id++;
      recipientData.push(`rId${id}`);
    }
    if (recipientData.length > 0) {
      ctx._settingsOptions = {
        ...settings,
        mailMerge: { ...mailMerge, odso: { ...odso, recipientData } },
      };
    }
  }
  const xmlifiedFileMapping = xmlifyContext(ctx);
  const files = compileMapping(xmlifiedFileMapping, overrides);
  const textBoxParts = compileTextBoxParts(ctx);
  for (const { part, relationships } of textBoxParts) {
    files[part.path] = typeof part.data === "string" ? encoder.encode(part.data) : part.data;
    if (relationships)
      files[relationships.path] =
        typeof relationships.data === "string"
          ? encoder.encode(relationships.data)
          : relationships.data;
  }
  if (ctx._options.customizations) {
    files["word/customizations.xml"] = encoder.encode(
      XML_DECL + (customizationsDesc.stringify(ctx._options.customizations, ctx) ?? ""),
    );
  }

  // Media + OLE embedding binaries (word/media/*, word/embeddings/*)
  addModelBinaries(files, "word", ctx.media.array, ctx.embeddings.array, mediaLevel);

  // Font files — only fonts carrying binary data produce a .odttf part.
  // Round-tripped fonts (rawOdttf) keep their original obfuscated bytes.
  for (const font of ctx.fontTable.fontOptionsWithKey) {
    for (const [partName, part] of embeddedFontParts(font)) {
      if (part.data === undefined) continue;
      // Pack the font part under the URI-escaped path so the ZIP item name
      // equals the escaped rel Target byte-for-byte (readers resolve the
      // Target as a URI; an unescaped part name with spaces never matches).
      const filePath = encodeUriPath(
        part.odttfPath ?? defaultEmbeddedFontPath(font.name, partName),
      );
      files[filePath] = part.rawOdttf ? part.data : obfuscate(part.data, part.fontKey);
    }
  }
  for (const font of ctx.glossaryFontTable?.fontOptionsWithKey ?? []) {
    for (const [partName, part] of embeddedFontParts(font)) {
      if (part.data === undefined) continue;
      const filePath = encodeUriPath(
        part.odttfPath ?? defaultEmbeddedFontPath(font.name, partName, "glossary"),
      );
      files[filePath] = part.rawOdttf ? part.data : obfuscate(part.data, part.fontKey);
    }
  }

  // [Content_Types].xml is serialized last: parts register their media/fonts
  // during stringify (run by xmlifyContext above), so deriving it now sees the
  // complete set. Building it inside xmlifyContext's object literal evaluated
  // it before header/footer/font media was registered, leaving jpg/gif/odttf
  // without a covering Default.
  files["[Content_Types].xml"] = encoder.encode(
    finalizeContentTypes(
      files,
      {
        resolve: DOCX_CONTENT_TYPE_RESOLVER,
        mediaContentTypes: DOCX_MEDIA_CONTENT_TYPES,
        source: ctx._options.contentTypes,
        rawParts: ctx._options.rawParts,
        overrides: [
          ...ctx.altChunks.array.map((ac) => ({
            path: `word/${ac.path}`,
            contentType: ac.contentType ?? "application/xhtml+xml",
          })),
          ...ctx.subDocs.array.map((sd) => ({
            path: `word/${sd.path}`,
            contentType:
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
          })),
        ],
        forcedOverrides: [
          {
            path: mainPartPath,
            contentType: resolveMainPartContentType(
              "wordprocessing",
              ctx._options.contentTypes,
              mainPartPath,
              packageFormat.mainContentType,
            ),
          },
          // Only forced when the source declared it (or fresh compile): some
          // producers cover the theme part with the Default `xml` entry, and
          // re-declaring it as an Override drifts from the source table.
          ...(hasSourcePart(ctx._options, themePartName(ctx._options))
            ? [
                {
                  path: themePartName(ctx._options),
                  contentType: DOCX_THEME_CONTENT_TYPE,
                },
              ]
            : []),
          ...ctx.chartExes.array.map((chartEx, index) => ({
            path:
              chartEx.sourcePath ??
              `word/charts/chartEx${Number(chartEx.key.replace(/^chartEx_/, "")) || index + 1}.xml`,
            contentType:
              sourcePartContentType(
                ctx._options,
                chartEx.sourcePath ??
                  `word/charts/chartEx${Number(chartEx.key.replace(/^chartEx_/, "")) || index + 1}.xml`,
              ) ?? DOCX_CHART_EX_CONTENT_TYPE,
          })),
          ...(ctx.glossaryOptions
            ? [
                {
                  path: `word/${ctx.glossaryOptions.partName ?? "glossary/document.xml"}`,
                  contentType:
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml",
                },
                ...(ctx.glossaryOptions.numbering
                  ? [
                      {
                        path: `word/${
                          ctx.glossaryOptions.numberingPartName ?? "glossary/numbering.xml"
                        }`,
                        contentType:
                          "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
                      },
                    ]
                  : []),
                ...(ctx.glossaryOptions.settings
                  ? [
                      {
                        path: `word/${
                          ctx.glossaryOptions.settingsPartName ?? "glossary/settings.xml"
                        }`,
                        contentType:
                          "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
                      },
                    ]
                  : []),
                ...(ctx.glossaryOptions.styles &&
                canForcePartContentType(
                  ctx._options,
                  `word/${ctx.glossaryOptions.stylesPartName ?? "glossary/styles.xml"}`,
                  "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
                )
                  ? [
                      {
                        path: `word/${ctx.glossaryOptions.stylesPartName ?? "glossary/styles.xml"}`,
                        contentType:
                          "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
                      },
                    ]
                  : []),
                ...(ctx.glossaryOptions.webSettings
                  ? [
                      {
                        path: `word/${
                          ctx.glossaryOptions.webSettingsPartName ?? "glossary/webSettings.xml"
                        }`,
                        contentType:
                          "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
                      },
                    ]
                  : []),
                ...(ctx.glossaryOptions.fonts?.length
                  ? [
                      {
                        path: `word/${
                          ctx.glossaryOptions.fontTablePartName ?? "glossary/fontTable.xml"
                        }`,
                        contentType:
                          "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
                      },
                    ]
                  : []),
                ...(ctx.glossaryOptions.footnotes
                  ? [
                      {
                        path: `word/${
                          ctx.glossaryOptions.footnotesPartName ?? "glossary/footnotes.xml"
                        }`,
                        contentType:
                          "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
                      },
                    ]
                  : []),
                ...(ctx.glossaryOptions.endnotes
                  ? [
                      {
                        path: `word/${
                          ctx.glossaryOptions.endnotesPartName ?? "glossary/endnotes.xml"
                        }`,
                        contentType:
                          "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
                      },
                    ]
                  : []),
                ...ctx.glossaryHeaderParts.map((header) => ({
                  path: `word/${header.partName}`,
                  contentType:
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
                })),
                ...ctx.glossaryFooterParts.map((footer) => ({
                  path: `word/${footer.partName}`,
                  contentType:
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
                })),
              ]
            : []),
        ],
      },
      ctx,
    ),
  );

  // Guard: drop passthrough rels whose target part never made it into the
  // package (hand-authored input) — Office refuses to open dangling rels.
  dropDanglingPassthroughRels(
    files,
    ctx._options.passthroughRelationships,
    (relationship) =>
      relationship.source === "" &&
      (relationship.relationshipType.includes("/core-properties") ||
        relationship.relationshipType.includes("/extended-properties")),
  );
  useStrictRelationshipTypes(files, documentNamespaceDialect(ctx));
  renamePrimaryDocument(files, packageFormat.mainPartPath, mainPartPath);

  return files;
}

// ── Internal ──

/**
 * Complete mapping of all XML files in an OOXML document package.
 */
interface XmlifyedFileMapping {
  Document: XmlifyedFile;
  Styles?: XmlifyedFile;
  Properties?: XmlifyedFile;
  Numbering?: XmlifyedFile;
  NumberingRelationships?: XmlifyedFile;
  Relationships: XmlifyedFile;
  FileRelationships: XmlifyedFile;
  Headers: XmlifyedFile[];
  Footers: XmlifyedFile[];
  HeaderRelationships: XmlifyedFile[];
  FooterRelationships: XmlifyedFile[];
  CustomProperties?: XmlifyedFile;
  AppProperties?: XmlifyedFile;
  FootNotes?: XmlifyedFile;
  FootNotesRelationships?: XmlifyedFile;
  Endnotes?: XmlifyedFile;
  EndnotesRelationships?: XmlifyedFile;
  Settings?: XmlifyedFile;
  MailMergeRecipients?: XmlifyedFile[];
  Comments?: XmlifyedFile;
  CommentsRelationships?: XmlifyedFile;
  People?: XmlifyedFile;
  CommentsExtended?: XmlifyedFile;
  FontTable?: XmlifyedFile;
  FontTableRelationships?: XmlifyedFile;
  Bibliography?: XmlifyedFile;
  Charts?: XmlifyedFile[];
  ChartEx?: XmlifyedFile[];
  DiagramData?: XmlifyedFile[];
  DiagramLayout?: XmlifyedFile[];
  DiagramStyle?: XmlifyedFile[];
  DiagramColors?: XmlifyedFile[];
  DiagramDrawing?: XmlifyedFile[];
  AltChunks?: XmlifyedFile[];
  SubDocs?: XmlifyedFile[];
  Glossary?: XmlifyedFile;
  GlossaryRelationships?: XmlifyedFile;
  GlossaryNumbering?: XmlifyedFile;
  GlossaryNumberingRelationships?: XmlifyedFile;
  GlossaryStyles?: XmlifyedFile;
  GlossarySettings?: XmlifyedFile;
  GlossarySettingsRelationships?: XmlifyedFile;
  GlossaryWebSettings?: XmlifyedFile;
  GlossaryFontTable?: XmlifyedFile;
  GlossaryFontTableRelationships?: XmlifyedFile;
  GlossaryFootnotes?: XmlifyedFile;
  GlossaryFootnotesRelationships?: XmlifyedFile;
  GlossaryEndnotes?: XmlifyedFile;
  GlossaryEndnotesRelationships?: XmlifyedFile;
  GlossaryHeaders?: XmlifyedFile[];
  GlossaryHeaderRelationships?: XmlifyedFile[];
  GlossaryFooters?: XmlifyedFile[];
  GlossaryFooterRelationships?: XmlifyedFile[];
  WebSettings?: XmlifyedFile;
}

function xmlifyContext(ctx: DocxWriteContext): XmlifyedFileMapping {
  const mailMergeRecipients = ctx._options.mailMergeRecipients ?? [];
  const settingsXml = hasSourcePart(ctx._options, "word/settings.xml")
    ? XML_DECL + (settingsDesc.stringify(ctx._settingsOptions, ctx) ?? "")
    : "";
  const hasSettingsRelationships = (ctx._options.passthroughRelationships ?? []).some(
    (rel) => rel.source === "word/settings.xml",
  );
  const hasAppProperties =
    hasSourcePart(ctx._options, "docProps/app.xml") || ctx._options.appProperties !== undefined;
  const recipientData =
    ctx._settingsOptions.mailMerge?.odso?.recipientData?.slice(0, mailMergeRecipients.length) ?? [];
  const mailMerge = ctx._settingsOptions.mailMerge;
  const mkCtx = (viewWrapper: DocxContext["viewWrapper"] = ctx.document): DocxContext => {
    const bodyCtx: DocxContext = {
      fileData: ctx,
      file: ctx,
      viewWrapper,
      reproducible: ctx.reproducible,
      addRelationship: (type: string, target: string, mode?: string) =>
        ctx.addRelationship(type, target, mode),
      addMedia: (data: Uint8Array, type: string, fileName?: string) =>
        ctx.addMedia(data, type, fileName),
      addHyperlink: (key, target) => ctx.addHyperlink(key, target),
      stringifyChild(child) {
        return stringifyBodyChild(child, bodyCtx);
      },
    };
    return bodyCtx;
  };

  const docCtx = mkCtx(ctx.document);
  const documentXmlData = XML_DECL + stringifyDocumentXml(ctx, docCtx);
  // Sampled after stringify: hyperlinks/altChunks registered during body
  // stringification take sequential ids here, and the media/embedding offsets
  // below must skip them (same ordering as the footnote part).
  const documentRelationshipCount = ctx.document.relationships.nextRelationshipId;

  // Phases owned by the compile/ folder, invoked in the order their parts
  // stringified before (document body → notes → headers/footers → the
  // document rels), so per-part media registration keeps its sequencing.
  const notesParts = compileNotesParts(ctx, mkCtx);
  const headerFooterParts = compileHeaderFooterParts(ctx, mkCtx);
  const glossaryParts = compileGlossaryParts(ctx, mkCtx);
  const documentEntries = compileDocumentEntries(ctx, documentXmlData, documentRelationshipCount);

  return {
    ...(hasAppProperties
      ? {
          AppProperties: {
            data:
              XML_DECL + (appPropertiesDesc.stringify(ctx._options.appProperties ?? {}, ctx) ?? ""),
            path: "docProps/app.xml",
          },
        }
      : {}),
    ...notesParts,
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
            data: XML_DECL + (peopleDesc.stringify(ctx._options.people, docCtx) ?? ""),
            path: "word/people.xml",
          },
        }
      : {}),
    ...(ctx._options.commentsExtended?.length
      ? {
          CommentsExtended: {
            data:
              XML_DECL +
              (commentsExtendedDesc.stringify(ctx._options.commentsExtended, docCtx) ?? ""),
            path: "word/commentsExtended.xml",
          },
        }
      : {}),
    ...(ctx._options.commentsIds != null
      ? {
          CommentsIds: {
            data: XML_DECL + (commentsIdsDesc.stringify(ctx._options.commentsIds, docCtx) ?? ""),
            path: "word/commentsIds.xml",
          },
        }
      : {}),
    ...(ctx._options.commentsExtensible != null
      ? {
          CommentsExtensible: {
            data:
              XML_DECL +
              (commentsExtensibleDesc.stringify(ctx._options.commentsExtensible, docCtx) ?? ""),
            path: "word/commentsExtensible.xml",
          },
        }
      : {}),
    ...documentEntries,
    // Theme — a parsed source theme round-trips structurally; fresh output
    // emits the Office default. Theme fill media resolve against part rels.
    ...(() => {
      const themePath = themePartName(ctx._options);
      if (
        ctx._options.theme === undefined &&
        ctx._options.rawParts?.some((part) => part.path === themePath)
      ) {
        return {};
      }
      const rels = new Relationships();
      const themeXml = ctx._options.theme
        ? bindThemeMedia(
            buildThemeXml(ctx._options.theme, withThemeRelationships(ctx, rels)),
            ctx,
            rels,
          )
        : createThemeXml();
      const themeFile = themePath.split("/").at(-1)!;
      return {
        Theme: { data: XML_DECL + themeXml, path: themePath },
        ...(rels.relationshipCount > 0
          ? {
              ThemeRelationships: {
                data: XML_DECL + rels.serialize(),
                path: `word/theme/_rels/${themeFile}.rels`,
              },
            }
          : {}),
      };
    })(),
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
    ...headerFooterParts,
    ...(ctx.hasNumbering
      ? (() => {
          // Picture-bullet media resolves through numbering.xml's own rels —
          // same placeholder bridge as headers/footers, ids local to the part
          // (numbering.xml carries no other relationships).
          const numberingMedia = resolveNumberingMedia(ctx.numbering.serialize(ctx), ctx);
          const numberingRels = numberingMedia.rels;
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
    ...(() => {
      const corePropertiesPath = ctx._options.corePropertiesPath ?? "docProps/core.xml";
      return hasSourcePart(ctx._options, corePropertiesPath)
        ? {
            Properties: {
              data: XML_DECL + (corePropertiesDesc.stringify(ctx._options, ctx) ?? ""),
              path: corePropertiesPath,
            },
          }
        : {};
    })(),
    ...(hasSourcePart(ctx._options, "word/settings.xml")
      ? {
          Settings: {
            data: settingsXml,
            path: "word/settings.xml",
          },
        }
      : {}),
    ...(mailMergeRecipients.length > 0
      ? {
          MailMergeRecipients: mailMergeRecipients.map((part, index) => ({
            data: XML_DECL + (mailMergeRecipientsDesc.stringify(part, ctx) ?? ""),
            path: `word/${part.partName ?? `recipients${index + 1}.xml`}`,
          })),
        }
      : {}),
    ...(ctx._settingsOptions.attachedTemplate !== undefined ||
    recipientData.length > 0 ||
    hasSettingsRelationships ||
    mailMerge?.dataSource !== undefined ||
    mailMerge?.headerSource !== undefined ||
    mailMerge?.odso?.src !== undefined
      ? {
          SettingsRelationships: {
            data: (() => {
              const attachedTemplateId =
                ctx._settingsOptions.attachedTemplate === undefined
                  ? undefined
                  : attachedTemplateRelationshipId(
                      ctx._settingsOptions.attachedTemplate,
                      ctx._options.passthroughRelationships ?? [],
                      recipientData,
                    );
              const rels = new Relationships();
              const settingsSourceRels = (ctx._options.passthroughRelationships ?? []).filter(
                (rel) => rel.source === "word/settings.xml",
              );
              const settingsSourceRids = new Set<string>();
              for (const match of (settingsXml ?? "").matchAll(/\br:id="([^"]+)"/g)) {
                const id = match[1];
                if (id) settingsSourceRids.add(id);
              }
              if (ctx._settingsOptions.attachedTemplate !== undefined) {
                rels.claimSourceRel({
                  relationshipType: RELATIONSHIP_TYPES.attachedTemplate,
                  target: ctx._settingsOptions.attachedTemplate,
                  rId: attachedTemplateId!,
                  targetMode: TargetModeType.EXTERNAL,
                });
              }
              for (const rel of settingsSourceRels) {
                const isRecipientSource =
                  recipientData.includes(rel.rId) &&
                  rel.relationshipType === RELATIONSHIP_TYPES.recipientData;
                const ownsAttachedTemplateId =
                  attachedTemplateId !== undefined && rel.rId === attachedTemplateId;
                if (
                  settingsSourceRids.has(rel.rId) &&
                  !isRecipientSource &&
                  !ownsAttachedTemplateId
                ) {
                  rels.claimSourceRel(rel, { force: true });
                }
              }
              const attachedTemplateTarget = ctx._settingsOptions.attachedTemplate;
              recipientData.forEach((rId, index) => {
                const id = /^rId(\d+)$/.exec(rId)?.[1];
                if (id) {
                  rels.claimSourceRel({
                    relationshipType: RELATIONSHIP_TYPES.recipientData,
                    target: mailMergeRecipients[index]?.partName ?? `recipients${index + 1}.xml`,
                    rId,
                  });
                }
              });
              for (const rel of settingsSourceRels) {
                const recipientIndex = recipientData.indexOf(rel.rId);
                const canonicalRecipientTarget =
                  recipientIndex >= 0
                    ? (mailMergeRecipients[recipientIndex]?.partName ??
                      `recipients${recipientIndex + 1}.xml`)
                    : undefined;
                if (
                  (attachedTemplateTarget !== undefined &&
                    rel.relationshipType === RELATIONSHIP_TYPES.attachedTemplate) ||
                  (canonicalRecipientTarget !== undefined &&
                    rel.relationshipType === RELATIONSHIP_TYPES.recipientData &&
                    rels.hasRelationship(rel.relationshipType, canonicalRecipientTarget))
                ) {
                  continue;
                }
                rels.claimSourceRel(rel);
              }
              return XML_DECL + rels.serialize();
            })(),
            path: "word/_rels/settings.xml.rels",
          },
        }
      : {}),
    ...(hasSourcePart(ctx._options, "word/styles.xml")
      ? {
          Styles: {
            data: (() => {
              const xmlStyles = ctx.styles.serialize(documentNamespaceDialect(ctx));
              return replaceNumberingPlaceholders(xmlStyles, ctx.numbering.concreteNumbering);
            })(),
            path: "word/styles.xml",
          },
        }
      : {}),
    ...(ctx._options.stylesWithEffects
      ? {
          StylesWithEffects: {
            data: (() => {
              const xmlStyles = ctx.stylesWithEffects!.serialize(documentNamespaceDialect(ctx));
              return replaceNumberingPlaceholders(xmlStyles, ctx.numbering.concreteNumbering);
            })(),
            path: "word/stylesWithEffects.xml",
          },
        }
      : {}),
    ...(ctx._options.bibliography
      ? {
          Bibliography: {
            data: XML_DECL + (bibliographyDesc.stringify(ctx._options.bibliography, ctx) ?? ""),
            path: "word/bibliography.xml",
          },
        }
      : {}),
    ...compileChartParts(ctx),
    ...compileChartExParts(ctx),
    ...compileSmartArtParts(ctx),
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
    ...glossaryParts.entries,
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
