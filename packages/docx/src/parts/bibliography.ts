/**
 * Bibliography module for WordprocessingML documents.
 *
 * Provides types for bibliography sources.
 * XML generation is handled by the descriptor pipeline (bibliographyDesc).
 *
 * Reference: ISO/IEC 29500-4, shared-bibliography.xsd, CT_Sources, CT_SourceType
 *
 * @module
 */

/**
 * Options for a single bibliography source entry.
 *
 * Maps to CT_SourceType in the bibliography XSD schema (choice of 48 elements,
 * all optional). All fields are optional — include only the relevant ones for
 * each source type. `author` is plain text (semicolon-separated); the XSD's
 * structured CT_AuthorType roles collapse into it.
 *
 * @property sourceType - Source type element (ST_SourceType: Book, JournalArticle, …)
 * @property type - Subcategory string (XSD Type element)
 * @property title - Title of the work
 * @property author - Author names (plain text, semicolon-separated)
 * @property year - Publication year
 * @property month - Publication month
 * @property day - Publication day
 * @property bookTitle - Title of the book (for book sections, articles in collections)
 * @property journal - Journal name
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
  /** Source type element (ST_SourceType: Book, JournalArticle, …). */
  sourceType?: string;
  /** Subcategory string (XSD Type element). */
  type?: string;
  title?: string;
  author?: string;
  year?: string;
  month?: string;
  day?: string;
  bookTitle?: string;
  journal?: string;
  volume?: string;
  issue?: string;
  pages?: string;
  publisher?: string;
  city?: string;
  url?: string;
  edition?: string;
  institution?: string;
  abbreviatedCaseNumber?: string;
  albumTitle?: string;
  broadcaster?: string;
  broadcastTitle?: string;
  caseNumber?: string;
  chapterNumber?: string;
  comments?: string;
  conferenceName?: string;
  countryRegion?: string;
  court?: string;
  dayAccessed?: string;
  department?: string;
  distributor?: string;
  guid?: string;
  internetSiteTitle?: string;
  lcid?: string;
  medium?: string;
  monthAccessed?: string;
  numberVolumes?: string;
  patentNumber?: string;
  periodicalTitle?: string;
  productionCompany?: string;
  publicationTitle?: string;
  recordingNumber?: string;
  refOrder?: string;
  reporter?: string;
  shortTitle?: string;
  standardNumber?: string;
  stateProvince?: string;
  station?: string;
  tag?: string;
  theater?: string;
  thesisType?: string;
  version?: string;
  yearAccessed?: string;
}

/**
 * Options for creating a bibliography container.
 *
 * @property sources - Array of bibliography source entries
 * @property styleName - Bibliography style name (e.g., "APA", "Chicago", "IEEE")
 */
export interface BibliographyOptions {
  sources: SourceTypeOptions[];
  styleName?: string;
}

// ── Descriptor ──

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { escapeXml, findChild, textOf } from "@office-open/xml";

const SOURCE_FIELDS: readonly (readonly [string, keyof SourceTypeOptions])[] = [
  ["SourceType", "sourceType"],
  ["Type", "type"],
  ["Title", "title"],
  ["Author", "author"],
  ["Year", "year"],
  ["Month", "month"],
  ["Day", "day"],
  ["BookTitle", "bookTitle"],
  ["JournalName", "journal"],
  ["Volume", "volume"],
  ["Issue", "issue"],
  ["Pages", "pages"],
  ["Publisher", "publisher"],
  ["City", "city"],
  ["URL", "url"],
  ["Edition", "edition"],
  ["Institution", "institution"],
  ["AbbreviatedCaseNumber", "abbreviatedCaseNumber"],
  ["AlbumTitle", "albumTitle"],
  ["Broadcaster", "broadcaster"],
  ["BroadcastTitle", "broadcastTitle"],
  ["CaseNumber", "caseNumber"],
  ["ChapterNumber", "chapterNumber"],
  ["Comments", "comments"],
  ["ConferenceName", "conferenceName"],
  ["CountryRegion", "countryRegion"],
  ["Court", "court"],
  ["DayAccessed", "dayAccessed"],
  ["Department", "department"],
  ["Distributor", "distributor"],
  ["Guid", "guid"],
  ["InternetSiteTitle", "internetSiteTitle"],
  ["LCID", "lcid"],
  ["Medium", "medium"],
  ["MonthAccessed", "monthAccessed"],
  ["NumberVolumes", "numberVolumes"],
  ["PatentNumber", "patentNumber"],
  ["PeriodicalTitle", "periodicalTitle"],
  ["ProductionCompany", "productionCompany"],
  ["PublicationTitle", "publicationTitle"],
  ["RecordingNumber", "recordingNumber"],
  ["RefOrder", "refOrder"],
  ["Reporter", "reporter"],
  ["ShortTitle", "shortTitle"],
  ["StandardNumber", "standardNumber"],
  ["StateProvince", "stateProvince"],
  ["Station", "station"],
  ["Tag", "tag"],
  ["Theater", "theater"],
  ["ThesisType", "thesisType"],

  ["Version", "version"],
  ["YearAccessed", "yearAccessed"],
] as const;

export const bibliographyDesc: CustomDescriptor<BibliographyOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const attrParts: string[] = [
      'xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"',
    ];
    if (opts.styleName !== undefined) {
      attrParts.push(`StyleName="${escapeXml(opts.styleName)}"`);
    }

    const parts: string[] = [`<b:Sources ${attrParts.join(" ")}>`];

    for (const source of opts.sources) {
      const sourceParts: string[] = [];
      for (const [tagName, key] of SOURCE_FIELDS) {
        const value = source[key];
        if (value !== undefined) {
          sourceParts.push(`<b:${tagName}>${escapeXml(value)}</b:${tagName}>`);
        }
      }
      parts.push(`<b:Source>${sourceParts.join("")}</b:Source>`);
    }

    parts.push("</b:Sources>");
    return parts.join("");
  },

  parse(el, _ctx) {
    const opts: Partial<BibliographyOptions> = {};

    // StyleName attribute
    const styleName = el.attributes?.["StyleName"];
    if (styleName) opts.styleName = styleName as string;

    // Parse b:Source children
    const sources: SourceTypeOptions[] = [];
    for (const child of el.elements ?? []) {
      if (child.name !== "b:Source") continue;
      const source: SourceTypeOptions = {};
      for (const [tagName, key] of SOURCE_FIELDS) {
        const xmlChild = findChild(child, `b:${tagName}`);
        if (xmlChild) {
          const val = textOf(xmlChild);
          if (val) source[key] = val;
        }
      }
      sources.push(source);
    }
    opts.sources = sources;

    return opts as BibliographyOptions;
  },
};
