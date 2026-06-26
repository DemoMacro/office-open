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
import type { SettingsOptions } from "@parts/settings/settings";
import { Styles, extractStyleId } from "@parts/styles";
import { ExternalStylesFactory } from "@parts/styles/external-styles-factory";
import { DefaultStylesFactory, collectDefaultOverrideIds } from "@parts/styles/factory";
import { SubDocCollection } from "@parts/sub-doc/sub-doc-collection";
import type { WebSettingsOptions } from "@parts/web-settings";
import { EmbeddingCollection } from "@shared/embeddings/embeddings";
import { Media } from "@shared/media";
import type { MediaData } from "@shared/media/data";
import type { SectionOptions } from "@shared/section";
import type { SectionChild } from "@shared/section";

import type { DocxDocument } from "./parse";

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
  declare public media: Media;
  declare public charts: ChartCollection;
  declare public smartArts: SmartArtCollection;
  declare public embeddings: EmbeddingCollection;
  declare public altChunks: AltChunkCollection;
  declare public subDocs: SubDocCollection;
  declare public comments: { relationships: Relationships };
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
    const fileName = this.media.nextMediaName(type);
    this.media.addImage(fileName, {
      data,
      fileName,
      type,
      transformation: { pixels: { x: 0, y: 0 }, emus: { x: 0, y: 0 } },
    } as MediaData);
    return `{${fileName}}`;
  }

  // --- Internal tracking ---
  private _headers: HeaderFooterEntry[] = [];
  private _footers: HeaderFooterEntry[] = [];

  // --- Original input preserved for descriptor usage ---
  declare public _options: DocumentOptions;

  constructor(options: DocumentOptions) {
    this._options = options;

    this.numbering = new Numbering(options.numbering ? options.numbering : { config: [] });

    this.comments = { relationships: new Relationships() };
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

    this.media = new Media();
    this.charts = new ChartCollection();
    this.smartArts = new SmartArtCollection();
    this.embeddings = new EmbeddingCollection();
    this.altChunks = new AltChunkCollection();
    this.subDocs = new SubDocCollection();

    if (options.externalStyles !== undefined) {
      const defaultFactory = new DefaultStylesFactory();
      const defaultStyles = defaultFactory.newInstance(options.styles?.default);
      const externalFactory = new ExternalStylesFactory();
      const externalStyles = externalFactory.newInstance(options.externalStyles);
      // Skip docDefaults AND latentStyles from default factory —
      // external styles already provide them; XSD requires docDefaults → latentStyles → style sequence.
      // Also drop builtins whose styleId the external XML already defines so
      // user definitions override defaults (no duplicate styleId in the output).
      const externalStyleIds = new Set<string>();
      for (const s of externalStyles.importedStyles ?? []) {
        const id = extractStyleId(s._raw);
        if (id) externalStyleIds.add(id);
      }
      const defaultStyleElements = defaultStyles.importedStyles!.slice(2).filter((s) => {
        const id = extractStyleId(s._raw);
        return !id || !externalStyleIds.has(id);
      });
      this.styles = new Styles({
        ...externalStyles,
        importedStyles: [...externalStyles.importedStyles!, ...defaultStyleElements],
      });
    } else if (options.styles) {
      const stylesFactory = new DefaultStylesFactory();
      const defaultStyles = stylesFactory.newInstance(options.styles.default);
      // importedStyles[0]=docDefaults, [1]=latentStyles, [2+]=builtin styles.
      // paragraphStyles/characterStyles stay available for numbering registration
      // below but are NOT emitted when round-tripping (the raw importedStyles
      // already carry every source style — emitting both would duplicate them).
      const {
        importedStyles: parsedStyles,
        paragraphStyles,
        characterStyles,
        tableStyles,
        ...restStyles
      } = options.styles;
      const merged = defaultStyles.importedStyles ? [...defaultStyles.importedStyles] : [];
      if (restStyles.docDefaultsXml) {
        const ddIdx = merged.findIndex((s) => s._raw.startsWith("<w:docDefaults"));
        if (ddIdx >= 0) merged[ddIdx] = { _raw: restStyles.docDefaultsXml };
      }
      if (restStyles.latentStylesXml) {
        const latentIdx = merged.findIndex((s) => s._raw.startsWith("<w:latentStyles"));
        if (latentIdx >= 0) merged[latentIdx] = { _raw: restStyles.latentStylesXml };
      }
      if (parsedStyles && parsedStyles.length > 0) {
        // Round-trip: emit verbatim source styles (suppress factory builtins
        // and structured re-emission). User overrides via default.<field>
        // (e.g. default.heading1) take precedence over the verbatim source —
        // drop those ids from the source and emit the factory's structured
        // version instead, so "modify a default style" works post-parse.
        merged.splice(2);
        const overrideIds = collectDefaultOverrideIds(options.styles.default);
        if (overrideIds.size > 0) {
          const overrideById = new Map<string, { _raw: string }>();
          for (const s of defaultStyles.importedStyles!.slice(2)) {
            const id = extractStyleId(s._raw);
            if (id) overrideById.set(id, s);
          }
          for (const s of parsedStyles) {
            const id = extractStyleId(s._raw);
            if (id && overrideIds.has(id)) continue;
            merged.push(s);
          }
          for (const id of overrideIds) {
            const s = overrideById.get(id);
            if (s) merged.push(s);
          }
        } else {
          merged.push(...parsedStyles);
        }
        this.styles = new Styles({ ...defaultStyles, importedStyles: merged, ...restStyles });
      } else {
        // Generation: factory builtins + structured custom styles.
        this.styles = new Styles({
          ...defaultStyles,
          paragraphStyles,
          characterStyles,
          tableStyles,
          ...restStyles,
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
    if (this._options.comments?.children?.length) {
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

    // Theme — a raw-passthrough part. Only declare the document→theme
    // relationship when the source carried one, so Word can resolve theme
    // colors/fonts. Without it theme1.xml is an orphan part Word may reject.
    const themePart = this._options.rawParts?.find((p) => p.path.startsWith("word/theme/"));
    if (themePart) {
      this.document.relationships.addRelationship(
        this._currentRelationshipId++,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
        themePart.path.replace(/^word\//, ""),
      );
    }

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
