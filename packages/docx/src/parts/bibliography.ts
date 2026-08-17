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
 * A person in a bibliography name list (CT_PersonType).
 */
export interface PersonOptions {
  last?: string;
  first?: string;
  middle?: string;
}

/**
 * A corporate (organization) author entry (b:Corporate) — only valid in the
 * author and performer roles, whose XSD type is CT_NameOrCorporateType.
 */
export interface CorporateOptions {
  corporate: string;
}

/** One entry of a role list: a person or (author/performer roles) a corporation. */
export type AuthorEntry = PersonOptions | CorporateOptions;

/**
 * Author container (CT_AuthorType) — one array per named role; each entry
 * serializes to one `b:<Role>` element. All roles are optional; include only
 * the ones relevant to the source type.
 */
export interface AuthorOptions {
  /** b:Author role */
  authors?: AuthorEntry[];
  /** b:Artist role */
  artists?: PersonOptions[];
  /** b:BookAuthor role */
  bookAuthors?: PersonOptions[];
  /** b:Compiler role */
  compilers?: PersonOptions[];
  /** b:Composer role */
  composers?: PersonOptions[];
  /** b:Conductor role */
  conductors?: PersonOptions[];
  /** b:Counsel role */
  counsel?: PersonOptions[];
  /** b:Director role */
  directors?: PersonOptions[];
  /** b:Editor role */
  editors?: PersonOptions[];
  /** b:Interviewee role */
  interviewees?: PersonOptions[];
  /** b:Interviewer role */
  interviewers?: PersonOptions[];
  /** b:Inventor role */
  inventors?: PersonOptions[];
  /** b:Performer role */
  performers?: AuthorEntry[];
  /** b:ProducerName role */
  producers?: PersonOptions[];
  /** b:Translator role */
  translators?: PersonOptions[];
  /** b:Writer role */
  writers?: PersonOptions[];
}

/**
 * Options for a single bibliography source entry.
 *
 * Maps to CT_SourceType in the bibliography XSD schema (choice of 48 elements,
 * all optional). All fields are optional — include only the relevant ones for
 * each source type. `author` maps the structured CT_AuthorType role container.
 *
 * @property sourceType - Source type element (ST_SourceType: Book, JournalArticle, …)
 * @property type - Subcategory string (XSD Type element)
 * @property title - Title of the work
 * @property author - Authors by role (CT_AuthorType)
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
  author?: AuthorOptions;
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
import type { Element as XmlElement } from "@office-open/xml";
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

// Role element tag → AuthorOptions key. Author/Performer are CT_NameOrCorporateType;
// every other role is CT_NameType (person list only).
const AUTHOR_ROLES: readonly (readonly [string, keyof AuthorOptions])[] = [
  ["Author", "authors"],
  ["Artist", "artists"],
  ["BookAuthor", "bookAuthors"],
  ["Compiler", "compilers"],
  ["Composer", "composers"],
  ["Conductor", "conductors"],
  ["Counsel", "counsel"],
  ["Director", "directors"],
  ["Editor", "editors"],
  ["Interviewee", "interviewees"],
  ["Interviewer", "interviewers"],
  ["Inventor", "inventors"],
  ["Performer", "performers"],
  ["ProducerName", "producers"],
  ["Translator", "translators"],
  ["Writer", "writers"],
] as const;

const ROLE_TAG_TO_KEY = new Map<string, keyof AuthorOptions>(
  AUTHOR_ROLES.map(([tag, key]) => [`b:${tag}`, key]),
);

/** Roles whose XSD type is CT_NameOrCorporateType (accept b:Corporate). */
const CORPORATE_ROLES = new Set(["b:Author", "b:Performer"]);

function stringifyAuthorEntry(entry: AuthorEntry): string {
  if ("corporate" in entry) {
    return `<b:Corporate>${escapeXml(entry.corporate)}</b:Corporate>`;
  }
  const parts: string[] = [];
  if (entry.last !== undefined) parts.push(`<b:Last>${escapeXml(entry.last)}</b:Last>`);
  if (entry.first !== undefined) parts.push(`<b:First>${escapeXml(entry.first)}</b:First>`);
  if (entry.middle !== undefined) parts.push(`<b:Middle>${escapeXml(entry.middle)}</b:Middle>`);
  return `<b:NameList><b:Person>${parts.join("")}</b:Person></b:NameList>`;
}

function stringifyAuthorElement(author: AuthorOptions): string {
  const roleParts: string[] = [];
  for (const [tagName, key] of AUTHOR_ROLES) {
    const entries = author[key];
    if (!entries) continue;
    for (const entry of entries) {
      roleParts.push(`<b:${tagName}>${stringifyAuthorEntry(entry)}</b:${tagName}>`);
    }
  }
  return roleParts.length > 0 ? `<b:Author>${roleParts.join("")}</b:Author>` : "";
}

function readPerson(el: XmlElement): PersonOptions {
  const person: PersonOptions = {};
  const last = findChild(el, "b:Last");
  if (last) {
    const val = textOf(last);
    if (val) person.last = val;
  }
  const first = findChild(el, "b:First");
  if (first) {
    const val = textOf(first);
    if (val) person.first = val;
  }
  const middle = findChild(el, "b:Middle");
  if (middle) {
    const val = textOf(middle);
    if (val) person.middle = val;
  }
  return person;
}

function parseAuthorElement(el: XmlElement): AuthorOptions | undefined {
  const author: { [K in keyof AuthorOptions]?: AuthorEntry[] } = {};
  for (const roleEl of el.elements ?? []) {
    const key = roleEl.name !== undefined ? ROLE_TAG_TO_KEY.get(roleEl.name) : undefined;
    if (!key) continue;

    const entries: AuthorEntry[] = [];
    const nameList = findChild(roleEl, "b:NameList");
    if (nameList) {
      for (const personEl of nameList.elements ?? []) {
        if (personEl.name !== "b:Person") continue;
        entries.push(readPerson(personEl));
      }
    } else if (roleEl.name !== undefined && CORPORATE_ROLES.has(roleEl.name)) {
      const corporate = findChild(roleEl, "b:Corporate");
      const name = corporate ? textOf(corporate) : undefined;
      if (name) entries.push({ corporate: name });
    }
    if (entries.length === 0) continue;
    author[key] = [...(author[key] ?? []), ...entries];
  }
  return Object.keys(author).length > 0 ? (author as AuthorOptions) : undefined;
}

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
        if (key === "author") {
          if (source.author !== undefined) {
            sourceParts.push(stringifyAuthorElement(source.author));
          }
          continue;
        }
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
        if (!xmlChild) continue;
        if (key === "author") {
          const parsed = parseAuthorElement(xmlChild);
          if (parsed) source.author = parsed;
          continue;
        }
        const val = textOf(xmlChild);
        if (val) source[key] = val;
      }
      sources.push(source);
    }
    opts.sources = sources;

    return opts as BibliographyOptions;
  },
};
