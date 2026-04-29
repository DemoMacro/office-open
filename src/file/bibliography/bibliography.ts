/**
 * Bibliography module for WordprocessingML documents.
 *
 * This module provides support for bibliography sources in documents.
 * Bibliography data is stored in a separate `word/bibliography.xml` part
 * within the DOCX package.
 *
 * Reference: ISO/IEC 29500-4, shared-bibliography.xsd, CT_Sources, CT_SourceType
 *
 * @module
 */
import { Relationships } from "@file/relationships";
import { XmlAttributeComponent, XmlComponent } from "@file/xml-components";

/**
 * Options for a single bibliography source entry.
 *
 * Maps to CT_SourceType in the bibliography XSD schema.
 * All fields are optional — include only the relevant ones for each source type.
 *
 * @property type - Source type (Book, JournalArticle, ConferenceProceedings, etc.)
 * @property title - Title of the work
 * @property author - Author names (plain text, semicolon-separated)
 * @property year - Publication year
 * @property month - Publication month
 * @property day - Publication day
 * @property bookTitle - Title of the book (for book sections, articles in collections)
 * @property journal - Journal name (JournalName in XSD)
 * @property volume - Volume number
 * @property issue - Issue number
 * @property pages - Page range
 * @property publisher - Publisher name
 * @property city - City of publication
 * @property url - URL for internet sources
 * @property edition - Edition number or description
 * @property institution - Institution (for theses, reports)
 */
export interface SourceTypeOptions {
    readonly type?: string;
    readonly title?: string;
    readonly author?: string;
    readonly year?: string;
    readonly month?: string;
    readonly day?: string;
    readonly bookTitle?: string;
    readonly journal?: string;
    readonly volume?: string;
    readonly issue?: string;
    readonly pages?: string;
    readonly publisher?: string;
    readonly city?: string;
    readonly url?: string;
    readonly edition?: string;
    readonly institution?: string;
}

/**
 * Options for creating a bibliography container.
 *
 * @property sources - Array of bibliography source entries
 * @property styleName - Bibliography style name (e.g., "APA", "Chicago", "IEEE")
 */
export interface IBibliographyOptions {
    readonly sources: readonly SourceTypeOptions[];
    readonly styleName?: string;
}

/** @internal */
class SourcesAttributes extends XmlAttributeComponent<{
    readonly "xmlns:b"?: string;
    readonly SelectedStyle?: string;
    readonly StyleName?: string;
    readonly URI?: string;
}> {
    protected readonly xmlKeys = {
        SelectedStyle: "SelectedStyle",
        StyleName: "StyleName",
        URI: "URI",
        "xmlns:b": "xmlns:b",
    };
}

/**
 * XML element containing text content with bibliography namespace prefix.
 *
 * @internal
 */
class BibliographyString extends XmlComponent {
    public constructor(tagName: string, value: string) {
        super(`b:${tagName}`);
        this.root.push(value);
    }
}

/**
 * Represents a single bibliography source (CT_SourceType).
 *
 * Each source is a collection of named text fields corresponding to
 * the elements defined in shared-bibliography.xsd.
 *
 * @internal
 */
class Source extends XmlComponent {
    public constructor(options: SourceTypeOptions) {
        super("b:Source");

        const fields: [string, string | undefined][] = [
            ["SourceType", options.type],
            ["Title", options.title],
            ["Author", options.author],
            ["Year", options.year],
            ["Month", options.month],
            ["Day", options.day],
            ["BookTitle", options.bookTitle],
            ["JournalName", options.journal],
            ["Volume", options.volume],
            ["Issue", options.issue],
            ["Pages", options.pages],
            ["Publisher", options.publisher],
            ["City", options.city],
            ["URL", options.url],
            ["Edition", options.edition],
            ["Institution", options.institution],
        ];

        for (const [tagName, value] of fields) {
            if (value !== undefined) {
                this.root.push(new BibliographyString(tagName, value));
            }
        }
    }
}

/**
 * Represents the bibliography sources container (CT_Sources).
 *
 * This is the root element for the `word/bibliography.xml` part that stores
 * all bibliography source definitions in the document.
 *
 * Reference: ISO/IEC 29500-4, shared-bibliography.xsd
 *
 * ## XSD Schema
 * ```xml
 * <xsd:element name="Sources" type="CT_Sources"/>
 * <xsd:complexType name="CT_Sources">
 *   <xsd:sequence>
 *     <xsd:element name="Source" type="CT_SourceType" minOccurs="0" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="SelectedStyle" type="s:ST_String"/>
 *   <xsd:attribute name="StyleName" type="s:ST_String"/>
 *   <xsd:attribute name="URI" type="s:ST_String"/>
 * </xsd:complexType>
 * ```
 *
 * @publicApi
 *
 * @example
 * ```typescript
 * new Bibliography({
 *   styleName: "APA",
 *   sources: [
 *     {
 *       type: "Book",
 *       title: "The Design of Everyday Things",
 *       author: "Norman, Donald",
 *       year: "2013",
 *       publisher: "Basic Books",
 *       city: "New York",
 *       edition: "Revised",
 *     },
 *   ],
 * });
 * ```
 */
export class Bibliography extends XmlComponent {
    private readonly relationships: Relationships;

    public constructor(options: IBibliographyOptions) {
        super("b:Sources");

        this.root.push(
            new SourcesAttributes({
                "xmlns:b": "http://purl.oclc.org/ooxml/officeDocument/bibliography",
                StyleName: options.styleName,
            }),
        );

        for (const source of options.sources) {
            this.root.push(new Source(source));
        }

        this.relationships = new Relationships();
    }

    public get Relationships(): Relationships {
        return this.relationships;
    }
}
