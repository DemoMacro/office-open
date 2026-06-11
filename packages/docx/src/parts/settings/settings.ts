/**
 * Settings module for WordprocessingML documents.
 *
 * This module provides document-level settings including compatibility,
 * track changes, headers/footers, and hyphenation options.
 *
 * Reference: http://officeopenxml.com/WPsettings.php
 *
 * @module
 */
import {
  FootnotePositionType,
  EndnotePositionType,
} from "@parts/document/body/section-properties/properties/footnote-endnote-properties";

import type { CompatibilityOptions } from "./compatibility";
export type { CompatibilityOptions } from "./compatibility";

/**
 * Options for configuring document settings.
 *
 * @see {@link Settings}
 */
export interface SettingsOptions {
  /** @deprecated Use compatibility.version instead */
  compatibilityModeVersion?: number;
  /** Enable different headers/footers for even and odd pages */
  evenAndOddHeaders?: boolean;
  /** Enable track changes (revision marking) */
  trackRevisions?: boolean;
  /** Do not track formatting changes when trackRevisions is on */
  doNotTrackFormatting?: boolean;
  /** Do not track move changes when trackRevisions is on */
  doNotTrackMoves?: boolean;
  /** Controls which types of revisions are visible */
  revisionView?: RevisionViewOptions;
  /** Update fields when document is opened */
  updateFields?: boolean;
  /** Compatibility settings for older Word versions */
  compatibility?: CompatibilityOptions;
  /** Default distance between tab stops in twips */
  defaultTabStop?: number;
  /** Hyphenation settings */
  hyphenation?: HyphenationOptions;
  /** Controls whether punctuation is compressed at line ends */
  characterSpacingControl?: "compressPunctuation" | "doNotCompress";
  /** Document protection settings */
  documentProtection?: DocumentProtectionOptions;
  /** Default document view mode */
  view?: "none" | "print" | "outline" | "masterPages" | "normal" | "web";
  /** Default zoom level (percentage) and type */
  zoom?: {
    percent?: number;
    val?: "none" | "fullPage" | "bestFit" | "textFit";
  };
  /** Write protection recommendation (not enforcement) */
  writeProtection?: WriteProtectionOptions;
  /** Whether to display the background shape in print layout */
  displayBackgroundShape?: boolean;
  /** Whether to display page boundaries between pages */
  doNotDisplayPageBoundaries?: boolean;
  /** Whether to embed TrueType fonts in the document */
  embedTrueTypeFonts?: boolean;
  /** Whether to embed system fonts in the document */
  embedSystemFonts?: boolean;
  /** Whether to save only a subset of the embedded fonts */
  saveSubsetFonts?: boolean;
  /** Document variables (key-value pairs stored in the document) */
  docVars?: { name: string; val: string }[];
  /** Mail merge configuration */
  mailMerge?: MailMergeOptions;
  /** Theme color scheme remapping */
  colorSchemeMapping?: {
    bg1?: string;
    t1?: string;
    bg2?: string;
    t2?: string;
    accent1?: string;
    accent2?: string;
    accent3?: string;
    accent4?: string;
    accent5?: string;
    accent6?: string;
    hyperlink?: string;
    followedHyperlink?: string;
  };
  /** Relationship ID to the attached template document */
  attachedTemplate?: string;
  /** Theme font language (BCP-47 tag, e.g. "en-US", "ja-JP") */
  themeFontLang?: string;
  /** Hide spelling errors in the document */
  hideSpellingErrors?: boolean;
  /** Hide grammatical errors in the document */
  hideGrammaticalErrors?: boolean;
  /** Disable punctuation kerning (CJK) */
  noPunctuationKerning?: boolean;
  /** Remove personal information when saving */
  removePersonalInformation?: boolean;
  /** Remove date and time metadata when saving */
  removeDateAndTime?: boolean;
  /** Print PostScript codes over text */
  printPostScriptOverText?: boolean;
  /** Print using fractional character widths */
  printFractionalCharacterWidth?: boolean;
  /** Print only form field data */
  printFormsData?: boolean;
  /** Save only form field data */
  saveFormsData?: boolean;
  /** Use mirror margins for facing pages */
  mirrorMargins?: boolean;
  /** Align document borders and edges with page edges */
  alignBordersAndEdges?: boolean;
  /** Page borders do not surround header content */
  bordersDoNotSurroundHeader?: boolean;
  /** Page borders do not surround footer content */
  bordersDoNotSurroundFooter?: boolean;
  /** Position gutter at top of page */
  gutterAtTop?: boolean;
  /** Document is in forms design mode */
  formsDesign?: boolean;
  /** Link styles from attached template */
  linkStyles?: boolean;
  /** Allow auto-format overrides */
  autoFormatOverride?: boolean;
  /** Lock document theme styles */
  styleLockTheme?: boolean;
  /** Lock quick format style set */
  styleLockQFSet?: boolean;
  /** Show envelope content in the document */
  showEnvelope?: boolean;
  /** Print two pages on one sheet */
  printTwoOnOne?: boolean;
  /** Enforce strict first and last character rules (CJK) */
  strictFirstAndLastChars?: boolean;
  /** Save a preview picture in the document */
  savePreviewPicture?: boolean;
  /** Do not validate custom XML against schema */
  doNotValidateAgainstSchema?: boolean;
  /** Save invalid XML markup */
  saveInvalidXml?: boolean;
  /** Ignore mixed content in custom XML */
  ignoreMixedContent?: boolean;
  /** Always show placeholder text for custom XML */
  alwaysShowPlaceholderText?: boolean;
  /** Do not demarcate invalid XML regions */
  doNotDemarcateInvalidXml?: boolean;
  /** Save only XML data (no formatting) */
  saveXmlDataOnly?: boolean;
  /** Use XSLT when saving */
  useXSLTWhenSaving?: boolean;
  /** Do not embed smart tags */
  doNotEmbedSmartTags?: boolean;
  /** Do not auto-compress pictures */
  doNotAutoCompressPictures?: boolean;
  /** Do not include subdocuments in word count */
  doNotIncludeSubdocsInStats?: boolean;
  /** Enable book fold printing */
  bookFoldPrinting?: boolean;
  /** Enable book fold reverse printing */
  bookFoldRevPrinting?: boolean;
  /** Default table style name */
  defaultTableStyle?: string;
  /** Decimal symbol for numeric fields */
  decimalSymbol?: string;
  /** List separator character */
  listSeparator?: string;
  /** Click and type paragraph style name */
  clickAndTypeStyle?: string;
  /** Summary length percentage (0-100) */
  summaryLength?: number;
  /** Number of sheets per booklet in book fold printing */
  bookFoldPrintingSheets?: number;
  /** Horizontal spacing for the drawing grid (twips) */
  drawingGridHorizontalSpacing?: number;
  /** Vertical spacing for the drawing grid (twips) */
  drawingGridVerticalSpacing?: number;
  /** Display horizontal gridlines every N units */
  displayHorizontalDrawingGridEvery?: number;
  /** Display vertical gridlines every N units */
  displayVerticalDrawingGridEvery?: number;
  /** Horizontal origin for the drawing grid (twips) */
  drawingGridHorizontalOrigin?: number;
  /** Vertical origin for the drawing grid (twips) */
  drawingGridVerticalOrigin?: number;
  /** Document-level footnote properties (CT_FtnDocProps) */
  footnotePr?: FootnotePropertiesOptions;
  /** Document-level endnote properties (CT_EdnDocProps) */
  endnotePr?: EndnotePropertiesOptions;
  /** Document revision save IDs (CT_DocRsids) */
  rsids?: RsidsOptions;
  /** Reading mode ink lock-down settings */
  readModeInkLockDown?: ReadModeInkLockDownOptions;
  /** Caption configuration (CT_Captions) */
  captions?: CaptionsOptions;
  /** Math properties (m:mathPr) */
  mathPr?: MathPropertiesOptions;
  /** Active writing style checking language/grammar settings */
  activeWritingStyle?: {
    lang?: string;
    vendorID?: string;
    dllVersion?: string;
    nlCheck?: boolean;
    checkStyle?: boolean;
    appCheck?: string;
    appName?: string;
  }[];
  /** Proofing state (spelling/grammar check status) */
  proofState?: {
    spelling?: "clean" | "dirty";
    grammar?: "clean" | "dirty";
  };
  /** Style pane format filter (which styles to show) */
  stylePaneFormatFilter?: {
    allStyles?: boolean;
    customStyles?: boolean;
    stylesInUse?: boolean;
    headingStyles?: boolean;
    numberingStyles?: boolean;
    tableStyles?: boolean;
    directFormattingOnRuns?: boolean;
    directFormattingOnParagraphs?: boolean;
    directFormattingOnNumbering?: boolean;
    directFormattingOnTables?: boolean;
    clearFormatting?: boolean;
    top3HeadingStyles?: boolean;
    visibleStyles?: boolean;
    alternateStyleNames?: boolean;
  };
  /** Style pane sort method */
  stylePaneSortMethod?: "name" | "priority" | "default" | "font";
  /** Document type classification */
  documentType?: "letter" | "eMail" | "notSpecified";
  /** Do not use margins for drawing grid origin */
  doNotUseMarginsForDrawingGridOrigin?: boolean;
  /** Do not shade form data fields */
  doNotShadeFormData?: boolean;
  /** Custom kinsoku line break characters after which line breaks are not allowed */
  noLineBreaksAfter?: { lang?: string; val?: string };
  /** Custom kinsoku line break characters before which line breaks are not allowed */
  noLineBreaksBefore?: { lang?: string; val?: string };
  /** Save through XSLT transform */
  saveThroughXslt?: {
    id?: string;
    val?: string;
    solutionID?: string;
  };
  /** Show XML tags in document */
  showXMLTags?: boolean;
  /** Always merge empty namespace */
  alwaysMergeEmptyNamespace?: boolean;
  /** Header shape defaults */
  hdrShapeDefaults?: boolean;
  /** Attached schema references */
  attachedSchema?: string[];
  /** Force schema upgrade */
  forceUpgrade?: boolean;
  /** Smart tag type definitions */
  smartTagType?: {
    namespace?: string;
    namespaceuri?: string;
    name?: string;
    url?: string;
  }[];
  /** Shape defaults */
  shapeDefaults?: boolean;
}

/**
 * Controls which types of revisions are visible in the document.
 */
export interface RevisionViewOptions {
  /** Show markup for insertions */
  markup?: boolean;
  /** Show comments */
  comments?: boolean;
  /** Show insertions and deletions */
  insDel?: boolean;
  /** Show formatting changes */
  formatting?: boolean;
  /** Show ink annotations */
  inkAnnotations?: boolean;
}

/**
 * Options for document protection (restrict editing).
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_DocProtect
 */
export interface DocumentProtectionOptions {
  /** Type of editing restriction */
  edit?: "none" | "readOnly" | "comments" | "trackedChanges" | "forms";
  /** Whether formatting is restricted */
  formatting?: boolean;
  /** Plaintext password — automatically hashed to hashValue/saltValue when provided */
  password?: string;
  /** Password hash (SHA-512 base64) */
  hashValue?: string;
  /** Password salt (base64) */
  saltValue?: string;
  /** Legacy password hash (Transitional XSD: w:hash) */
  hash?: string;
  /** Legacy password salt (Transitional XSD: w:salt) */
  salt?: string;
  /** Password spin count */
  spinCount?: number;
  /** Password algorithm name */
  algorithmName?: string;
  /** Cryptographic algorithm class */
  cryptoAlgorithmClass?: string;
  /** Cryptographic algorithm SID */
  cryptoAlgorithmSid?: number;
  /** Cryptographic algorithm type */
  cryptoAlgorithmType?: string;
  /** Cryptographic provider */
  cryptoProvider?: string;
  /** Cryptographic provider type */
  cryptoProviderType?: string;
  /** Cryptographic provider type extension */
  cryptoProviderTypeExtension?: number;
  /** Cryptographic provider type extension source */
  cryptoProviderTypeExtensionSource?: string;
  /** Algorithm extension ID */
  algorithmExtensionId?: number;
  /** Algorithm extension source */
  algorithmExtensionSource?: string;
  /** Legacy cryptographic spin count (AG_TransitionalPassword) */
  cryptoSpinCount?: number;
}

/**
 * Options for write protection (read-only recommendation, not enforcement).
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_WriteProtection
 */
export interface WriteProtectionOptions {
  /** Plaintext password — automatically hashed to hashValue/saltValue when provided */
  password?: string;
  /** Cryptographic hash of the password */
  hashValue?: string;
  /** Salt value for the hash (base64) */
  saltValue?: string;
  /** Legacy password hash (Transitional XSD: w:hash) */
  hash?: string;
  /** Legacy password salt (Transitional XSD: w:salt) */
  salt?: string;
  /** Password spin count */
  spinCount?: number;
  /** Password algorithm name */
  algorithmName?: string;
  /** Whether write protection is recommended (default true when options provided) */
  recommended?: boolean;
  /** Cryptographic algorithm class */
  cryptoAlgorithmClass?: string;
  /** Cryptographic algorithm SID */
  cryptoAlgorithmSid?: number;
  /** Cryptographic algorithm type */
  cryptoAlgorithmType?: string;
  /** Cryptographic provider */
  cryptoProvider?: string;
  /** Cryptographic provider type */
  cryptoProviderType?: string;
  /** Cryptographic provider type extension */
  cryptoProviderTypeExtension?: number;
  /** Cryptographic provider type extension source */
  cryptoProviderTypeExtensionSource?: string;
  /** Algorithm extension ID */
  algorithmExtensionId?: number;
  /** Algorithm extension source */
  algorithmExtensionSource?: string;
  /** Legacy cryptographic spin count (AG_TransitionalPassword) */
  cryptoSpinCount?: number;
}

// ── Mail Merge types ──

/** Mail merge main document type (ST_MailMergeDocType) */
export type MailMergeDocType =
  | "catalog"
  | "envelopes"
  | "mailingLabels"
  | "formLetters"
  | "email"
  | "fax";

/** Mail merge destination (ST_MailMergeDest) */
export type MailMergeDest = "newDocument" | "printer" | "email" | "fax";

/** Mail merge data source type (ST_MailMergeDataType) */
export type MailMergeDataType =
  | "textFile"
  | "database"
  | "spreadsheet"
  | "email"
  | "odbc"
  | "native"
  | "addressBook"
  | "legacy"
  | "master";

/** Mail merge source type for ODSO (ST_MailMergeSourceType) */
export type MailMergeSourceType =
  | "database"
  | "addressBook"
  | "document1"
  | "document2"
  | "text"
  | "email"
  | "native"
  | "legacy"
  | "master";

/** ODSO field map type (ST_MailMergeOdsoFMDFieldType) */
export type OdsoFieldType = "null" | "dbColumn";

/** Field mapping for ODSO (CT_OdsoFieldMapData) */
export interface OdsoFieldMapDataOptions {
  type?: OdsoFieldType;
  name?: string;
  mappedName?: string;
  column?: number;
  lid?: string;
  dynamicAddress?: boolean;
}

/** Office Data Source Object (CT_Odso) */
export interface OdsoOptions {
  udl?: string;
  table?: string;
  src?: string;
  colDelim?: number;
  type?: MailMergeSourceType;
  fHdr?: boolean;
  fieldMapData?: OdsoFieldMapDataOptions[];
  recipientData?: string[];
  /** Unique tag for identifying the data source (w:uniqueTag) */
  uniqueTag?: string;
}

/** Mail merge configuration (CT_MailMerge) */
export interface MailMergeOptions {
  /** Main document type (required) */
  mainDocumentType: MailMergeDocType;
  /** Data source type (required) */
  dataType: MailMergeDataType;
  /** Destination for merged documents */
  destination?: MailMergeDest;
  /** Database connection string */
  connectString?: string;
  /** SQL query to select data */
  query?: string;
  /** Path to data source (relationship ID) */
  dataSource?: string;
  /** Path to header source (relationship ID) */
  headerSource?: string;
  /** Do not suppress blank lines */
  doNotSuppressBlankLines?: boolean;
  /** Address field name for email merge */
  addressFieldName?: string;
  /** Email subject line */
  mailSubject?: string;
  /** Send as email attachment */
  mailAsAttachment?: boolean;
  /** View merged data in document */
  viewMergedData?: boolean;
  /** Active record index */
  activeRecord?: number;
  /** Check errors mode */
  checkErrors?: number;
  /** Link to query in data source */
  linkToQuery?: boolean;
  /** Office Data Source Object configuration */
  odso?: OdsoOptions;
  /** Whether this mail merge is the active one (w:active) */
  active?: boolean;
  /** Recipients data reference (w:recipients r:id) */
  recipients?: string;
}

/**
 * Options for automatic hyphenation settings.
 *
 * @see {@link Settings}
 */
export interface HyphenationOptions {
  /** Specifies whether the application automatically hyphenates words as they are typed in the document. */
  autoHyphenation?: boolean;
  /** Specifies the minimum number of characters at the beginning of a word before a hyphen can be inserted. */
  hyphenationZone?: number;
  /** Specifies the maximum number of consecutive lines that can end with a hyphenated word. */
  consecutiveHyphenLimit?: number;
  /** Specifies whether to hyphenate words in all capital letters. */
  doNotHyphenateCaps?: boolean;
}

/**
 * Footnote properties for document-level settings (CT_FtnDocProps).
 */
export interface FootnotePropertiesOptions {
  /** Footnote placement */
  pos?: (typeof FootnotePositionType)[keyof typeof FootnotePositionType];
  /** Number format */
  numFmt?: string;
  /** Custom number format string */
  format?: string;
  /** Starting number */
  numStart?: number;
  /** When to restart numbering */
  numRestart?: string;
}

/**
 * Endnote properties for document-level settings (CT_EdnDocProps).
 */
export interface EndnotePropertiesOptions {
  /** Endnote placement */
  pos?: (typeof EndnotePositionType)[keyof typeof EndnotePositionType];
  /** Number format */
  numFmt?: string;
  /** Custom number format string */
  format?: string;
  /** Starting number */
  numStart?: number;
  /** When to restart numbering */
  numRestart?: string;
}

/** Document revision save IDs (CT_DocRsids) */
export interface RsidsOptions {
  /** Root revision save ID (8 hex characters) */
  rsidRoot?: string;
  /** List of revision save IDs */
  rsids?: string[];
}

/** Reading mode ink lock-down (CT_ReadingModeInkLockDown) */
export interface ReadModeInkLockDownOptions {
  /** Use actual page dimensions */
  actualPg?: boolean;
  /** Page width in pixels */
  w: number;
  /** Page height in pixels */
  h: number;
  /** Font size (percentage or points) */
  fontSz: number;
}

/** Caption definition (CT_Caption) */
export interface CaptionOptions {
  /** Caption style name */
  name: string;
  /** Caption position */
  pos?: "above" | "below";
  /** Include chapter number */
  chapNum?: boolean;
  /** Heading level for chapter number */
  heading?: number;
  /** Exclude label */
  noLabel?: boolean;
  /** Number format */
  numFmt?: string;
  /** Chapter separator */
  sep?: "hyphen" | "period" | "colon" | "emDash" | "enDash";
}

/** Auto-caption (CT_AutoCaption) */
export interface AutoCaptionOptions {
  /** Object type name */
  name: string;
  /** Caption style name to apply */
  caption: string;
}

/** Captions configuration (CT_Captions) */
export interface CaptionsOptions {
  /** Caption definitions */
  captions: CaptionOptions[];
  /** Auto-caption definitions */
  autoCaptions?: AutoCaptionOptions[];
}

/** Math properties (CT_MathPr) */
export interface MathPropertiesOptions {
  /** Default math font */
  mathFont?: string;
  /** Binary operator break style ("before" | "after" | "repeat") */
  brkBin?: string;
  /** Subtraction binary operator break ("--" | "-+" | "+-") */
  brkBinSub?: string;
  /** Use small fractions */
  smallFrac?: boolean;
  /** Use display defaults */
  dispDef?: boolean;
  /** Left margin (twips) */
  lMargin?: number;
  /** Right margin (twips) */
  rMargin?: number;
  /** Default justification ("centerGroup" | "center") */
  defJc?: string;
  /** Wrap indent (twips) */
  wrapIndent?: number;
  /** Integral limit location ("subSup" | "undOvr") */
  intLim?: string;
  /** N-ary limit location ("subSup" | "undOvr") */
  naryLim?: string;
}
