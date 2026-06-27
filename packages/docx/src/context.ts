/**
 * DOCX compilation context.
 *
 * DocxWriteContext holds all mutable state needed during document compilation.
 * generateDocument() creates a DocxWriteContext internally.
 *
 * @module
 */

import { Relationships } from "@office-open/core";
import { ChartCollection } from "@office-open/core/chart";
import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { SmartArtCollection } from "@office-open/core/smartart";
import type { Element } from "@office-open/xml";
import { AltChunkCollection } from "@parts/alt-chunk/alt-chunk-collection";
import type { DocumentOptions } from "@parts/core-properties";
import type { SectionPropertiesOptions } from "@parts/document/body/section-properties/section-properties";
import type { EndnoteSeparator } from "@parts/endnotes/descriptor";
import { FontWrapper } from "@parts/fonts/font-wrapper";
import type { FootnoteSeparator } from "@parts/footnotes/descriptor";
import type { GlossaryDocumentOptions } from "@parts/glossary-document";
import type { HeaderFooterEntry } from "@parts/header-footer";
import { Numbering } from "@parts/numbering";
import type { ParagraphOptions } from "@parts/paragraph/paragraph";
import type { CommentOptions } from "@parts/paragraph/run/comment-run";
import type { SettingsOptions } from "@parts/settings/settings";
import { Styles, extractStyleId } from "@parts/styles";
import { ExternalStylesFactory } from "@parts/styles/external-styles-factory";
import { DefaultStylesFactory } from "@parts/styles/factory";
import { SubDocCollection } from "@parts/sub-doc/sub-doc-collection";
import type { WebSettingsOptions } from "@parts/web-settings";
import { EmbeddingCollection } from "@shared/embeddings/embeddings";
import { Media } from "@shared/media";
import type { MediaData } from "@shared/media/data";
import type { SectionOptions } from "@shared/section";
import type { SectionChild } from "@shared/section";

import type { DocxDocument } from "./parse";

/** User styles override factory defaults with the same styleId; keep the rest. */
function mergeById<T extends { id: string }>(
  factoryStyles: T[] | undefined,
  userStyles: T[] | undefined,
): T[] {
  const factory = factoryStyles ?? [];
  if (!userStyles || userStyles.length === 0) return factory;
  const userIds = new Set(userStyles.map((s) => s.id));
  return [...factory.filter((s) => !userIds.has(s.id)), ...userStyles];
}

/**
 * Highest comment id in an explicit comments list, or -1 when there are none.
 * Seeds the comment id allocator so auto-allocated ids never collide with ids
 * the caller already assigned (e.g. round-tripped from an existing document).
 */
function maxCommentId(comments: readonly CommentOptions[] | undefined): number {
  let max = -1;
  if (comments) {
    for (const c of comments) {
      if (c.id > max) max = c.id;
    }
  }
  return max;
}

/**
 * Whether any `{ comment }` sugar child appears anywhere in the body tree
 * (paragraphs, tables, textboxes, SDTs, headers/footers nested in sections).
 * The document→comments relationship must exist whenever comments.xml will be
 * generated; since sugar entries are registered during stringify — after the
 * constructor wires relationships — this pre-scan predicts them so the part and
 * its relationship stay in sync (OPC consistency). Every `{ comment }` always
 * stringifies, so the prediction matches the entries actually registered.
 */
function bodyContainsCommentSugar(value: unknown): boolean {
  if (value === null || value === undefined) return false;
  if (typeof value !== "object") return false;
  if (value instanceof Uint8Array || value instanceof Date) return false;
  if (Array.isArray(value)) {
    for (const item of value) {
      if (bodyContainsCommentSugar(item)) return true;
    }
    return false;
  }
  const obj = value as Record<string, unknown>;
  if (typeof obj.comment === "object" && obj.comment !== null) return true;
  for (const key of Object.keys(obj)) {
    if (bodyContainsCommentSugar(obj[key])) return true;
  }
  return false;
}

/** Interface for document view wrappers — provides relationships access. */
export interface ViewWrapper {
  relationships: Relationships;
}

// ── BodyContext ──

/**
 * Context for body-level stringification.
 *
 * Pure JSON pipeline context — extends WriteContext for descriptor compatibility.
 * No dependency on XmlComponent Context (compile/ uses zero toXml calls).
 */
export interface BodyContext extends WriteContext {
  /** The root write context with all mutable document state. */
  fileData: DocxWriteContext;
  /** Alias for fileData — some descriptor internals access context.file. */
  file: DocxWriteContext;
  /** Current view wrapper for relationship access. */
  viewWrapper: { relationships: Relationships };
  /** Stringify a body-level child element — injected to break circular imports. */
  stringifyChild: (child: SectionChild, ctx: BodyContext) => string;
}

// ── DocxWriteContext ──

export class DocxWriteContext implements WriteContext {
  private _currentRelationshipId = 1;

  // --- Accessed by XmlComponent via context.file.* during toXml() ---
  declare public document: { relationships: Relationships };
  declare public numbering: Numbering;
  declare public media: Media<MediaData>;
  declare public charts: ChartCollection;
  declare public smartArts: SmartArtCollection;
  declare public embeddings: EmbeddingCollection;
  declare public altChunks: AltChunkCollection;
  declare public subDocs: SubDocCollection;
  declare public comments: {
    relationships: Relationships;
    /** Comment entries registered by `{ comment }` sugar children during stringify. */
    entries: CommentOptions[];
    /** Next auto-allocated comment id (seeded above any explicit comment id). */
    nextId: number;
  };
  declare public footNotes: {
    relationships: Relationships;
    notes: Map<number, (ParagraphOptions | string)[]>;
    separator?: FootnoteSeparator;
    continuationSeparator?: FootnoteSeparator;
  };
  declare public endnotes: {
    relationships: Relationships;
    notes: Map<number, (ParagraphOptions | string)[]>;
    separator?: EndnoteSeparator;
    continuationSeparator?: EndnoteSeparator;
  };

  // --- Additional state used by the compiler ---
  declare public fileRelationships: Relationships;
  declare public _settingsOptions: SettingsOptions;
  declare public styles: Styles;
  declare public fontTable: FontWrapper;
  declare public glossaryOptions: GlossaryDocumentOptions | undefined;
  declare public webSettings: WebSettingsOptions | undefined;

  // --- Section properties (one per section, raw options for descriptor pipeline) ---
  private _sectionProperties: SectionPropertiesOptions[] = [];
  public get sectionProperties(): readonly SectionPropertiesOptions[] {
    return this._sectionProperties;
  }

  // --- WriteContext interface (core descriptor pipeline) ---

  public addRelationship(_type: string, _target: string, _mode?: string): string {
    const id = this._currentRelationshipId++;
    return `rId${id}`;
  }

  public addMedia(data: Uint8Array, type: string): string {
    const entry = this.media.addMedia(
      data,
      type,
      (fileName) =>
        ({
          data,
          fileName,
          type,
          transformation: { pixels: { x: 0, y: 0 }, emus: { x: 0, y: 0 } },
        }) as MediaData,
    );
    return `{${entry.fileName}}`;
  }

  // --- Internal tracking ---
  private _headers: HeaderFooterEntry[] = [];
  private _footers: HeaderFooterEntry[] = [];

  // --- Original input preserved for descriptor usage ---
  declare public _options: DocumentOptions;

  constructor(options: DocumentOptions) {
    this._options = options;

    this.numbering = new Numbering(options.numbering ? options.numbering : { config: [] });

    this.comments = {
      relationships: new Relationships(),
      entries: [],
      nextId: maxCommentId(options.comments?.children) + 1,
    };
    this.fileRelationships = new Relationships();
    this.footNotes = { relationships: new Relationships(), notes: new Map() };
    this.endnotes = { relationships: new Relationships(), notes: new Map() };
    this.document = { relationships: new Relationships() };
    this._settingsOptions = {
      compatibility: options.compatibility,
      compatibilityModeVersion: options.compatabilityModeVersion,
      defaultTabStop: options.defaultTabStop,
      evenAndOddHeaders: options.evenAndOddHeaderAndFooters ? true : false,
      characterSpacingControl: options.characterSpacingControl,
      hyphenation: {
        autoHyphenation: options.hyphenation?.autoHyphenation,
        consecutiveHyphenLimit: options.hyphenation?.consecutiveHyphenLimit,
        doNotHyphenateCaps: options.hyphenation?.doNotHyphenateCaps,
        hyphenationZone: options.hyphenation?.hyphenationZone,
      },
      trackRevisions: options.features?.trackRevisions,
      updateFields: options.features?.updateFields,
      documentProtection: options.features?.documentProtection,
      view: options.view,
      zoom: options.zoom,
      writeProtection: options.writeProtection,
      displayBackgroundShape:
        options.displayBackgroundShape ?? (options.background?.image ? true : undefined),
      embedTrueTypeFonts: options.embedTrueTypeFonts,
      embedSystemFonts: options.embedSystemFonts,
      saveSubsetFonts: options.saveSubsetFonts,
      docVars: options.docVars,
      colorSchemeMapping: options.colorSchemeMapping,
      mailMerge: options.mailMerge,
      ...options.settings,
    };

    this.media = new Media<MediaData>();
    this.charts = new ChartCollection();
    this.smartArts = new SmartArtCollection();
    this.embeddings = new EmbeddingCollection();
    this.altChunks = new AltChunkCollection();
    this.subDocs = new SubDocCollection();

    if (options.externalStyles !== undefined) {
      const externalStyles = new ExternalStylesFactory().newInstance(options.externalStyles);
      const defaultStyles = new DefaultStylesFactory().newInstance(options.styles?.default ?? {});
      // External (user-provided full styles.xml) wins; factory builtins fill
      // any gaps. Drop factory builtins whose styleId the external XML already
      // defines (no duplicate styleId). docDefaults/latentStyles come from the
      // external XML — the factory's are not mixed in.
      const externalIds = new Set<string>();
      for (const s of externalStyles.importedStyles ?? []) {
        const id = extractStyleId(s._raw);
        if (id) externalIds.add(id);
      }
      const notInExternal = <T extends { id: string }>(arr: T[] | undefined) =>
        (arr ?? []).filter((s) => !externalIds.has(s.id));
      this.styles = new Styles({
        importedStyles: externalStyles.importedStyles,
        initialAttributes: externalStyles.initialAttributes ?? defaultStyles.initialAttributes,
        paragraphStyles: notInExternal(defaultStyles.paragraphStyles),
        characterStyles: notInExternal(defaultStyles.characterStyles),
        tableStyles: notInExternal(defaultStyles.tableStyles),
        numberingStyles: notInExternal(defaultStyles.numberingStyles),
      });
    } else if (options.styles) {
      const s = options.styles;
      if (s.roundTripped) {
        // Round-trip origin (parseStyleDefinitions): parsed structured
        // builtin/custom styles win. The factory only supplies docDefaults +
        // latentStyles verbatim defaults; parsed docDefaultsXml/latentStylesXml
        // override when present. No factory builtin rebuild — parsed builtins
        // already carry the source document's customizations.
        const f = new DefaultStylesFactory().newInstance({});
        const docDefaults = s.docDefaultsXml ?? f.importedStyles?.[0]?._raw ?? "";
        const latentStyles = s.latentStylesXml ?? f.importedStyles?.[1]?._raw ?? "";
        this.styles = new Styles({
          importedStyles: [{ _raw: docDefaults }, { _raw: latentStyles }],
          initialAttributes: s.initialAttributes ?? f.initialAttributes,
          paragraphStyles: s.paragraphStyles,
          characterStyles: s.characterStyles,
          tableStyles: s.tableStyles,
          numberingStyles: s.numberingStyles,
        });
      } else {
        // Fresh generation: factory default builtins (structured) + user
        // overrides. User paragraphStyles/characterStyles/tableStyles/
        // numberingStyles override factory builtins with the same styleId.
        const f = new DefaultStylesFactory().newInstance(s.default);
        this.styles = new Styles({
          importedStyles: f.importedStyles,
          initialAttributes: s.initialAttributes ?? f.initialAttributes,
          paragraphStyles: mergeById(f.paragraphStyles, s.paragraphStyles),
          characterStyles: mergeById(f.characterStyles, s.characterStyles),
          tableStyles: mergeById(f.tableStyles, s.tableStyles),
          numberingStyles: mergeById(f.numberingStyles, s.numberingStyles),
        });
      }
    } else {
      const stylesFactory = new DefaultStylesFactory();
      this.styles = new Styles(stylesFactory.newInstance());
    }

    // Register numbering references from custom paragraph/character styles.
    // Style definitions may contain numbering properties whose concrete instances
    // are never created through the body paragraph processing path.
    if (options.styles?.paragraphStyles) {
      for (const style of options.styles.paragraphStyles) {
        const num = style.paragraph?.numbering;
        if (num) {
          this.numbering.createConcreteNumberingInstance(num.reference, num.instance ?? 0);
        }
      }
    }

    this.addDefaultRelationships();

    for (const section of options.sections) {
      this.addSection(section);
    }

    if (options.footnotes) {
      for (const key in options.footnotes) {
        // Skip the round-tripped separator markers (they carry no .children).
        if (key === "separator" || key === "continuationSeparator") continue;
        this.footNotes.notes.set(parseFloat(key), options.footnotes[key].children);
      }
      this.footNotes.separator = options.footnotes.separator;
      this.footNotes.continuationSeparator = options.footnotes.continuationSeparator;
    }

    if (options.endnotes) {
      for (const key in options.endnotes) {
        if (key === "separator" || key === "continuationSeparator") continue;
        this.endnotes.notes.set(parseFloat(key), options.endnotes[key].children);
      }
      this.endnotes.separator = options.endnotes.separator;
      this.endnotes.continuationSeparator = options.endnotes.continuationSeparator;
    }

    this.fontTable = new FontWrapper(options.fonts ?? []);
    this.glossaryOptions = options.glossary;
    this.webSettings = options.webSettings ?? undefined;

    if (options.glossary) {
      this.document.relationships.addRelationship(
        this._currentRelationshipId++,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument",
        "glossary/document.xml",
      );
    }

    if (this.webSettings) {
      this.document.relationships.addRelationship(
        this._currentRelationshipId++,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings",
        "webSettings.xml",
      );
    }
  }

  get headers(): HeaderFooterEntry[] {
    return this._headers;
  }

  get footers(): HeaderFooterEntry[] {
    return this._footers;
  }

  // --- Private helpers ---

  private addSection({ headers = {}, footers = {}, properties }: SectionOptions): void {
    const sectPrOptions: SectionPropertiesOptions = {
      ...properties,
      footerWrapperGroup: {
        default: footers.default ? this.createFooter(footers.default) : undefined,
        even: footers.even ? this.createFooter(footers.even) : undefined,
        first: footers.first ? this.createFooter(footers.first) : undefined,
      },
      headerWrapperGroup: {
        default: headers.default ? this.createHeader(headers.default) : undefined,
        even: headers.even ? this.createHeader(headers.even) : undefined,
        first: headers.first ? this.createHeader(headers.first) : undefined,
      },
    };
    this._sectionProperties.push(sectPrOptions);
  }

  private createHeader(header: SectionChild[]): HeaderFooterEntry {
    const referenceId = this._currentRelationshipId++;
    const entry: HeaderFooterEntry = {
      children: header,
      relationships: new Relationships(),
      referenceId,
    };
    this.addHeaderToDocument(entry);
    return entry;
  }

  private createFooter(footer: SectionChild[]): HeaderFooterEntry {
    const referenceId = this._currentRelationshipId++;
    const entry: HeaderFooterEntry = {
      children: footer,
      relationships: new Relationships(),
      referenceId,
    };
    this.addFooterToDocument(entry);
    return entry;
  }

  private addHeaderToDocument(header: HeaderFooterEntry): void {
    this._headers.push(header);
    this.document.relationships.addRelationship(
      header.referenceId,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
      `header${this._headers.length}.xml`,
    );
  }

  private addFooterToDocument(footer: HeaderFooterEntry): void {
    this._footers.push(footer);
    this.document.relationships.addRelationship(
      footer.referenceId,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
      `footer${this._footers.length}.xml`,
    );
  }

  private addDefaultRelationships(): void {
    this.fileRelationships.addRelationship(
      1,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
      "word/document.xml",
    );
    this.fileRelationships.addRelationship(
      2,
      "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
      "docProps/core.xml",
    );
    this.fileRelationships.addRelationship(
      3,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
      "docProps/app.xml",
    );
    this.fileRelationships.addRelationship(
      4,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
      "docProps/custom.xml",
    );

    this.document.relationships.addRelationship(
      this._currentRelationshipId++,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
      "styles.xml",
    );
    this.document.relationships.addRelationship(
      this._currentRelationshipId++,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
      "numbering.xml",
    );
    this.document.relationships.addRelationship(
      this._currentRelationshipId++,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
      "footnotes.xml",
    );
    this.document.relationships.addRelationship(
      this._currentRelationshipId++,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
      "endnotes.xml",
    );
    this.document.relationships.addRelationship(
      this._currentRelationshipId++,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
      "settings.xml",
    );
    // Comments is an optional part — only wire the document→comments relationship
    // when the document actually carries comments. Emitting it unconditionally
    // produces an orphan comments.xml that Word rejects as an OPC violation
    // (empty part with no [Content_Types] Override when content types are
    // passed through from the source on round-trip).
    if (
      this._options.comments?.children?.length ||
      bodyContainsCommentSugar(this._options.sections)
    ) {
      this.document.relationships.addRelationship(
        this._currentRelationshipId++,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
        "comments.xml",
      );
    }
    if (this._options.bibliography) {
      this.document.relationships.addRelationship(
        this._currentRelationshipId++,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/bibliography",
        "bibliography.xml",
      );
    }

    // Theme — always present: fresh-compile generates a default theme, round-trip
    // passes the source theme through rawParts. Word needs the document→theme
    // relationship to resolve theme colors/fonts.
    const themePart = this._options.rawParts?.find((p) => p.path.startsWith("word/theme/"));
    this.document.relationships.addRelationship(
      this._currentRelationshipId++,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
      themePart ? themePart.path.replace(/^word\//, "") : "theme/theme1.xml",
    );

    // customXml storage — raw-passthrough parts. Declare the document→customXml
    // relationship for each item (itemProps are linked via the item's own .rels,
    // not directly by the document) so Word can bind cover-page metadata, etc.
    for (const part of this._options.rawParts ?? []) {
      if (
        part.path.startsWith("customXml/") &&
        part.path.endsWith(".xml") &&
        !part.path.includes("/_rels/") &&
        !part.path.includes("itemProps")
      ) {
        this.document.relationships.addRelationship(
          this._currentRelationshipId++,
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
          `../${part.path}`,
        );
      }
    }
  }
}

// ── DocxReadContext ──

/**
 * DOCX-specific read context.
 *
 * Holds references to the parsed DocxDocument and cached style/numbering data
 * used throughout the DOCX parsing pipeline. Implements ReadContext for
 * descriptor pipeline compatibility.
 */
export class DocxReadContext implements ReadContext {
  /**
   * Path of the part currently being parsed. Each part carries its own .rels
   * with independent rId numbering, so drawings inside a part must resolve
   * image relationships against that part's rels. Defaults to the document body.
   */
  public currentPart = "word/document.xml";

  constructor(
    public docx: DocxDocument,
    public styleCache: Map<string, Element>,
    public numberingCache: Map<string, Element>,
  ) {}

  resolveRelationship(rId: string): string | undefined {
    const partMedia = this.docx.partRefs.partMedia.get(this.currentPart);
    if (partMedia) {
      const media = partMedia.get(rId);
      if (media) return media;
    }
    return (
      this.docx.partRefs.headers.get(rId) ??
      this.docx.partRefs.footers.get(rId) ??
      this.docx.partRefs.media.get(rId) ??
      this.docx.partRefs.charts.get(rId) ??
      this.docx.partRefs.diagramData.get(rId) ??
      this.docx.partRefs.afChunks.get(rId) ??
      this.docx.partRefs.subDocs.get(rId) ??
      this.docx.partRefs.hyperlinks.get(rId)
    );
  }

  /**
   * Run `fn` with `currentPart` temporarily set to `partPath`, restoring the
   * previous value afterwards. Use when parsing a sub-document part (header,
   * footer, footnotes, …) so its drawings resolve images from its own rels.
   */
  withPart<T>(partPath: string, fn: () => T): T {
    const prev = this.currentPart;
    this.currentPart = partPath;
    try {
      return fn();
    } finally {
      this.currentPart = prev;
    }
  }

  getPart(path: string): Element | undefined {
    return this.docx.doc.get(path);
  }

  getRaw(path: string): Uint8Array | undefined {
    return this.docx.doc.getRaw(path);
  }
}
