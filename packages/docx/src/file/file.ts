import { AltChunkCollection } from "./alt-chunk/alt-chunk-collection";
/**
 * File module for WordprocessingML documents.
 *
 * The File class is the main entry point for creating DOCX documents.
 * It manages all document parts including content, styles, numbering, and media.
 *
 * @module
 */
import { AppProperties } from "./app-properties/app-properties";
import { Bibliography } from "./bibliography";
import { ChartCollection } from "./chart/chart-collection";
import { coerceSectionChild } from "./coerce";
import { ContentTypes } from "./content-types/content-types";
import { CoreProperties } from "./core-properties";
import type { PropertiesOptions } from "./core-properties";
import { CustomProperties } from "./custom-properties";
import { DocumentWrapper } from "./document-wrapper";
import { HeaderFooterReferenceType } from "./document/body/section-properties";
import type { ISectionPropertiesOptions } from "./document/body/section-properties";
import { EndnotesWrapper } from "./endnotes-wrapper";
import { FontWrapper } from "./fonts/font-wrapper";
import { FooterWrapper } from "./footer-wrapper";
import type { DocumentFooter } from "./footer-wrapper";
import { FootnotesWrapper } from "./footnotes-wrapper";
import { Footer, Header } from "./header";
import { HeaderWrapper } from "./header-wrapper";
import type { DocumentHeader } from "./header-wrapper";
import { Media } from "./media";
import { Numbering } from "./numbering";
import { Paragraph } from "./paragraph/paragraph";
import { Comments } from "./paragraph/run/comment-run";
import { Relationships } from "./relationships";
import type { SectionChild } from "./section-child";
import { Settings } from "./settings";
import { SmartArtCollection } from "./smartart/smartart-collection";
import { Styles } from "./styles";
import { ExternalStylesFactory } from "./styles/external-styles-factory";
import { DefaultStylesFactory } from "./styles/factory";
import { SubDocCollection } from "./sub-doc/sub-doc-collection";

/**
 * Options for a document section.
 *
 * Each section can have its own headers, footers, and page properties.
 *
 * @property headers - Optional header definitions for the section
 * @property headers.default - Default header for all pages (when first/even not specified)
 * @property headers.first - Header for the first page of the section
 * @property headers.even - Header for even-numbered pages
 * @property footers - Optional footer definitions for the section
 * @property footers.default - Default footer for all pages (when first/even not specified)
 * @property footers.first - Footer for the first page of the section
 * @property footers.even - Footer for even-numbered pages
 * @property properties - Section properties such as page size, margins, and orientation
 * @property children - Array of content elements (paragraphs, tables, etc.) for this section
 */
export interface SectionOptions {
  /** Optional header definitions for the section. */
  readonly headers?: {
    /** Default header for all pages (when first/even not specified). */
    readonly default?: Header | readonly SectionChild[];
    /** Header for the first page of the section. */
    readonly first?: Header | readonly SectionChild[];
    /** Header for even-numbered pages. */
    readonly even?: Header | readonly SectionChild[];
  };
  /** Optional footer definitions for the section. */
  readonly footers?: {
    /** Default footer for all pages (when first/even not specified). */
    readonly default?: Footer | readonly SectionChild[];
    /** Footer for the first page of the section. */
    readonly first?: Footer | readonly SectionChild[];
    /** Footer for even-numbered pages. */
    readonly even?: Footer | readonly SectionChild[];
  };
  /** Section properties such as page size, margins, and orientation. */
  readonly properties?: ISectionPropertiesOptions;
  /** Array of content elements (paragraphs, tables, etc.) for this section. */
  readonly children: readonly SectionChild[];
}

/**
 * Represents a Word document file.
 *
 * The File class (exported as `Document`) is the main entry point for creating DOCX documents.
 * It manages all document components including content, styles, numbering, headers/footers,
 * and media. Documents are organized into sections, each of which can have its own page
 * settings, headers, and footers.
 *
 * This class handles the assembly of all OOXML parts required for a valid .docx file,
 * including relationships, content types, and document properties.
 *
 * @publicApi
 *
 * @example
 * ```typescript
 * // Simple document with one section
 * const doc = new Document({
 *   sections: [{
 *     children: [
 *       new Paragraph("Hello World"),
 *     ],
 *   }],
 * });
 *
 * // Document with multiple sections and headers/footers
 * const doc = new Document({
 *   creator: "John Doe",
 *   sections: [
 *     {
 *       headers: {
 *         default: new Header({
 *           children: [new Paragraph("Header Text")],
 *         }),
 *       },
 *       children: [
 *         new Paragraph("Section 1 content"),
 *       ],
 *     },
 *     {
 *       children: [
 *         new Paragraph("Section 2 content"),
 *       ],
 *     },
 *   ],
 * });
 *
 * // Document with custom styles and numbering
 * const doc = new Document({
 *   styles: {
 *     paragraphStyles: [
 *       {
 *         id: "MyHeading",
 *         name: "My Heading",
 *         basedOn: "Heading1",
 *         run: { bold: true, color: "FF0000" },
 *       },
 *     ],
 *   },
 *   numbering: {
 *     config: [
 *       {
 *         reference: "my-numbering",
 *         levels: [
 *           { level: 0, format: "decimal", text: "%1.", alignment: "left" },
 *         ],
 *       },
 *     ],
 *   },
 *   sections: [{
 *     children: [new Paragraph("Content")],
 *   }],
 * });
 * ```
 */
export class File {
  private _currentRelationshipId: number = 1;

  public readonly document: DocumentWrapper;
  private readonly _headers: DocumentHeader[] = [];
  private readonly _footers: DocumentFooter[] = [];
  public readonly coreProperties: CoreProperties;
  public readonly numbering: Numbering;
  public readonly media: Media;
  public readonly charts: ChartCollection;
  public readonly smartArts: SmartArtCollection;
  public readonly altChunks: AltChunkCollection;
  public readonly subDocs: SubDocCollection;
  public readonly fileRelationships: Relationships;
  public readonly footNotes: FootnotesWrapper;
  public readonly endnotes: EndnotesWrapper;
  public readonly settings: Settings;
  public readonly contentTypes: ContentTypes;
  public readonly customProperties: CustomProperties;
  public readonly appProperties: AppProperties;
  public readonly styles: Styles;
  public readonly comments: Comments;
  public readonly bibliography: Bibliography | undefined;
  public readonly fontTable: FontWrapper;

  public constructor(options: PropertiesOptions) {
    this.coreProperties = new CoreProperties({
      ...options,
      creator: options.creator ?? "Un-named",
      lastModifiedBy: options.lastModifiedBy ?? "Un-named",
      revision: options.revision ?? 1,
    });

    this.numbering = new Numbering(options.numbering ? options.numbering : { config: [] });

    this.comments = new Comments(options.comments ?? { children: [] });
    this.bibliography = options.bibliography ? new Bibliography(options.bibliography) : undefined;
    this.fileRelationships = new Relationships();
    this.customProperties = new CustomProperties(options.customProperties ?? []);
    this.appProperties = new AppProperties();
    this.footNotes = new FootnotesWrapper();
    this.endnotes = new EndnotesWrapper();
    this.contentTypes = new ContentTypes();
    this.document = new DocumentWrapper({ background: options.background });
    this.settings = new Settings({
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
    });

    this.media = new Media();
    this.charts = new ChartCollection();
    this.smartArts = new SmartArtCollection();
    this.altChunks = new AltChunkCollection();
    this.subDocs = new SubDocCollection();

    if (options.externalStyles !== undefined) {
      const defaultFactory = new DefaultStylesFactory();
      const defaultStyles = defaultFactory.newInstance(options.styles?.default);
      const externalFactory = new ExternalStylesFactory();
      const externalStyles = externalFactory.newInstance(options.externalStyles);
      // External styles already have docDefaults + latentStyles + styles in XSD order.
      // Append only the style elements from default factory (skip its DocumentDefaults
      // to avoid a second docDefaults after style elements, which violates XSD sequence).
      const defaultStyleElements = defaultStyles.importedStyles!.slice(1);
      this.styles = new Styles({
        ...externalStyles,
        importedStyles: [...externalStyles.importedStyles!, ...defaultStyleElements],
      });
    } else if (options.styles) {
      const stylesFactory = new DefaultStylesFactory();
      const defaultStyles = stylesFactory.newInstance(options.styles.default);
      this.styles = new Styles({
        ...defaultStyles,
        ...options.styles,
      });
    } else {
      const stylesFactory = new DefaultStylesFactory();
      this.styles = new Styles(stylesFactory.newInstance());
    }

    this.addDefaultRelationships();

    if (this.bibliography) {
      this.contentTypes.addBibliography();
    }

    for (const section of options.sections) {
      this.addSection(section);
    }

    if (options.footnotes) {
      for (const key in options.footnotes) {
        const children = options.footnotes[key].children.map((p) =>
          p instanceof Paragraph ? p : new Paragraph(p),
        );
        this.footNotes.view.createFootNote(parseFloat(key), children);
      }
    }

    if (options.endnotes) {
      for (const key in options.endnotes) {
        const children = options.endnotes[key].children.map((p) =>
          p instanceof Paragraph ? p : new Paragraph(p),
        );
        this.endnotes.view.createEndnote(parseFloat(key), children);
      }
    }

    this.fontTable = new FontWrapper(options.fonts ?? []);
  }

  private addSection({ headers = {}, footers = {}, children, properties }: SectionOptions): void {
    this.document.view.body.addSection({
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
    });

    for (const rawChild of children) {
      this.document.view.add(coerceSectionChild(rawChild));
    }
  }

  private createHeader(header: Header | readonly SectionChild[]): HeaderWrapper {
    const wrapper = new HeaderWrapper(this.media, this._currentRelationshipId++);

    const children = header instanceof Header ? header.options.children : header;
    for (const rawChild of children) {
      wrapper.add(coerceSectionChild(rawChild));
    }

    this.addHeaderToDocument(wrapper);
    return wrapper;
  }

  private createFooter(footer: Footer | readonly SectionChild[]): FooterWrapper {
    const wrapper = new FooterWrapper(this.media, this._currentRelationshipId++);

    const children = footer instanceof Footer ? footer.options.children : footer;
    for (const rawChild of children) {
      wrapper.add(coerceSectionChild(rawChild));
    }

    this.addFooterToDocument(wrapper);
    return wrapper;
  }

  private addHeaderToDocument(
    header: HeaderWrapper,
    type: (typeof HeaderFooterReferenceType)[keyof typeof HeaderFooterReferenceType] = HeaderFooterReferenceType.DEFAULT,
  ): void {
    this._headers.push({ header, type });
    this.document.relationships.addRelationship(
      header.view.referenceId,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
      `header${this._headers.length}.xml`,
    );
    this.contentTypes.addHeader(this._headers.length);
  }

  private addFooterToDocument(
    footer: FooterWrapper,
    type: (typeof HeaderFooterReferenceType)[keyof typeof HeaderFooterReferenceType] = HeaderFooterReferenceType.DEFAULT,
  ): void {
    this._footers.push({ footer, type });
    this.document.relationships.addRelationship(
      footer.view.referenceId,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
      `footer${this._footers.length}.xml`,
    );
    this.contentTypes.addFooter(this._footers.length);
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
    this.document.relationships.addRelationship(
      this._currentRelationshipId++,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
      "comments.xml",
    );
    if (this.bibliography) {
      this.document.relationships.addRelationship(
        this._currentRelationshipId++,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/bibliography",
        "bibliography.xml",
      );
    }
  }

  public get headers(): readonly HeaderWrapper[] {
    return this._headers.map((item) => item.header);
  }

  public get footers(): readonly FooterWrapper[] {
    return this._footers.map((item) => item.footer);
  }
}
