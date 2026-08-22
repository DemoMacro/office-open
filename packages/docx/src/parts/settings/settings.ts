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
import type { VmlShapeDefaultsOptions, VmlShapeLayoutOptions } from "@office-open/core";
import type { ColorSchemeIndex } from "@office-open/core";
import { NumberRestartType } from "@parts/document/body/section-properties/properties/footnote-endnote-properties";
import type { NumberFormat } from "@shared/constants";

import type { CompatibilityOptions } from "./compatibility";
export type { CompatibilityOptions, CompatSettingOptions } from "./compatibility";

/**
 * Shape defaults content (CT_ShapeDefaults) — the o: element sequence hosted
 * by `<w:hdrShapeDefaults>` and `<w:shapeDefaults>`.
 */
export interface ShapeDefaultsOptions {
  /** o:shapedefaults — VML shape defaults (fill/stroke/textbox/colormru/…). */
  shapedefaults?: VmlShapeDefaultsOptions;
  /** o:shapelayout — shape id map / regroup table / rules. */
  shapelayout?: VmlShapeLayoutOptions;
}

/**
 * Options for configuring document settings.
 *
 * @see {@link Settings}
 */
export interface SettingsOptions {
  /**
   * Verbatim inner XML of `<w:settings>` (all child elements). When set (from
   * parse), generate emits it verbatim so the full CT_Settings content — compat
   * flags, math properties, rsids, footnote/endnote properties, shape defaults,
   * clrSchemeMapping, etc. (~100 element types, most without a structured API) —
   * round-trips byte-for-byte. Delete this to fall back to structured generation.
   */
  rawXml?: string;
  /**
   * Root `<w:settings>` attributes captured verbatim from the source (`xmlns:*`
   * declarations + `mc:Ignorable`). Preserves source-specific namespaces (e.g.
   * `xmlns:sl`, `xmlns:wpsCustomData`) that the fixed SETTINGS_NS constant
   * omits, so rawXml child elements using those prefixes stay well-formed.
   */
  rootAttributes?: Record<string, string>;
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
  /**
   * Compatibility settings for older Word versions.
   *
   * Tri-state: omit (undefined) to emit the MS Office default compatSettings
   * for fresh documents; pass an object for explicit control; set `false` to
   * omit `<w:compat>` entirely.
   */
  compatibility?: CompatibilityOptions | false;
  /** Default distance between tab stops in twips */
  defaultTabStop?: number;
  /** Automatically hyphenate words as they are typed (w:autoHyphenation) */
  autoHyphenation?: boolean;
  /** Maximum number of consecutive lines ending with a hyphenated word (w:consecutiveHyphenLimit) */
  consecutiveHyphenLimit?: number;
  /** Distance from the margin within which hyphenation is avoided, in twips (w:hyphenationZone) */
  hyphenationZone?: number;
  /** Hyphenate words in all capital letters (w:doNotHyphenateCaps) */
  doNotHyphenateCaps?: boolean;
  /** Controls whether punctuation is compressed at line ends */
  characterSpacingControl?:
    | "compressPunctuation"
    | "compressPunctuationAndJapaneseKana"
    | "doNotCompress";
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
  /** Theme color scheme remapping (w:clrSchemeMapping, ST_WmlColorSchemeIndex) */
  colorSchemeMapping?: {
    bg1?: ColorSchemeIndex;
    t1?: ColorSchemeIndex;
    bg2?: ColorSchemeIndex;
    t2?: ColorSchemeIndex;
    accent1?: ColorSchemeIndex;
    accent2?: ColorSchemeIndex;
    accent3?: ColorSchemeIndex;
    accent4?: ColorSchemeIndex;
    accent5?: ColorSchemeIndex;
    accent6?: ColorSchemeIndex;
    hyperlink?: ColorSchemeIndex;
    followedHyperlink?: ColorSchemeIndex;
  };
  /** URL of the attached template document (external relationship target) */
  attachedTemplate?: string;
  /** Theme font languages (CT_Language): latin, eastAsian, and complex-script BCP-47 tags */
  themeFontLang?: { val?: string; eastAsia?: string; bidi?: string };
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
  useXSLTWhenSaving?: boolean;
  doNotEmbedSmartTags?: boolean;
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
  /** Word 2010 document identifier (w14:docId/`@w14:val`, e.g. "1A190769") */
  w14DocId?: string;
  /** Discard cropped-out image data when saving (w14:discardImageEditingData) */
  w14DiscardImageEditingData?: boolean;
  /** Default image DPI for pictures inserted in this document (w14:defaultImageDpi) */
  w14DefaultImageDpi?: number;
  /** Track chart references by document (w15:chartTrackingRefBased) */
  w15ChartTrackingRefBased?: boolean;
  /** Word 2013 document identifier (w15:docId/`@w15:val`, GUID format) */
  w15DocId?: string;
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
  /** Document-level footnote properties (CT_FtnDocProps, w:footnotePr) */
  footnoteProperties?: DocumentFootnotePropertiesOptions;
  /** Document-level endnote properties (CT_EdnDocProps, w:endnotePr) */
  endnoteProperties?: DocumentEndnotePropertiesOptions;
  /** Document revision save IDs (CT_DocRsids) */
  rsids?: RsidsOptions;
  /** Reading mode ink lock-down settings */
  readModeInkLockDown?: ReadModeInkLockDownOptions;
  /** Caption configuration (CT_Captions) */
  captions?: CaptionsOptions;
  /** Math properties (m:mathPr) */
  mathProperties?: MathPropertiesOptions;
  /** Emulate Word 97-2003 UI behavior (w:uiCompat97To2003, Word 2010+) */
  uiCompat97To2003?: boolean;
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
    latentStyles?: boolean;
    /** Legacy Word 2007 hex bitmask (ST_ShortHexNumber) — string keeps
     *  leading zeros; Word 2010+ writes the boolean attributes instead. */
    val?: string;
  };
  /** Style pane sort method (ST_StyleSort) */
  stylePaneSortMethod?:
    | "name"
    | "priority"
    | "default"
    | "font"
    | "basedOn"
    | "type"
    | "0000"
    | "0001"
    | "0002"
    | "0003"
    | "0004"
    | "0005";
  /** Document type classification */
  documentType?: "letter" | "eMail" | "notSpecified";
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
  alwaysMergeEmptyNamespace?: boolean;
  /** Header shape defaults (w:hdrShapeDefaults) */
  hdrShapeDefaults?: ShapeDefaultsOptions;
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
  /** Shape defaults (w:shapeDefaults) */
  shapeDefaults?: ShapeDefaultsOptions;
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
  /** Cryptographic algorithm class (w:cryptAlgorithmClass, s:ST_AlgClass) */
  cryptoAlgorithmClass?: "hash" | "custom";
  /** Cryptographic algorithm SID */
  cryptoAlgorithmSid?: number;
  /** Cryptographic algorithm type (w:cryptAlgorithmType, s:ST_AlgType) */
  cryptoAlgorithmType?: "typeAny" | "custom";
  /** Cryptographic provider */
  cryptoProvider?: string;
  /** Cryptographic provider type (w:cryptProviderType, s:ST_CryptProv) */
  cryptoProviderType?: "rsaAES" | "rsaFull" | "custom";
  /** Cryptographic provider type extension */
  cryptoProviderTypeExtension?: number;
  /** Cryptographic provider type extension source */
  cryptoProviderTypeExtensionSource?: string;
  algorithmExtensionId?: number;
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
  /** Cryptographic algorithm class (w:cryptAlgorithmClass, s:ST_AlgClass) */
  cryptoAlgorithmClass?: "hash" | "custom";
  /** Cryptographic algorithm SID */
  cryptoAlgorithmSid?: number;
  /** Cryptographic algorithm type (w:cryptAlgorithmType, s:ST_AlgType) */
  cryptoAlgorithmType?: "typeAny" | "custom";
  /** Cryptographic provider */
  cryptoProvider?: string;
  /** Cryptographic provider type (w:cryptProviderType, s:ST_CryptProv) */
  cryptoProviderType?: "rsaAES" | "rsaFull" | "custom";
  /** Cryptographic provider type extension */
  cryptoProviderTypeExtension?: number;
  /** Cryptographic provider type extension source */
  cryptoProviderTypeExtensionSource?: string;
  algorithmExtensionId?: number;
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
 * Footnote properties for document-level settings (CT_FtnDocProps).
 */
export interface DocumentFootnotePropertiesOptions {
  /** Footnote placement */
  pos?: "pageBottom" | "beneathText" | "sectEnd" | "docEnd";
  /** Number format (w:numFmt, ST_NumberFormat) */
  numFmt?: (typeof NumberFormat)[keyof typeof NumberFormat];
  /** Custom number format string */
  format?: string;
  /** Starting number */
  numStart?: number;
  /** When to restart numbering (w:numRestart, ST_RestartNumber) */
  numRestart?: (typeof NumberRestartType)[keyof typeof NumberRestartType];
  /**
   * Special footnotes acting as separator/continuation, as footnote ids in
   * word/footnotes.xml (Word writes the separator as -1 and the continuation
   * note as 0).
   */
  footnotes?: number[];
}

/**
 * Endnote properties for document-level settings (CT_EdnDocProps).
 */
export interface DocumentEndnotePropertiesOptions {
  /** Endnote placement */
  pos?: "sectEnd" | "docEnd";
  /** Number format (w:numFmt, ST_NumberFormat) */
  numFmt?: (typeof NumberFormat)[keyof typeof NumberFormat];
  /** Custom number format string */
  format?: string;
  /** Starting number */
  numStart?: number;
  /** When to restart numbering (w:numRestart, ST_RestartNumber) */
  numRestart?: (typeof NumberRestartType)[keyof typeof NumberRestartType];
  /**
   * Special endnotes acting as separator/continuation, as endnote ids in
   * word/endnotes.xml (Word writes the separator as -1 and the continuation
   * note as 0).
   */
  endnotes?: number[];
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
  /** Caption position (ST_CaptionPos) */
  pos?: "above" | "below" | "left" | "right";
  /** Include chapter number */
  chapterNumber?: boolean;
  /** Heading level for chapter number */
  heading?: number;
  /** Exclude label */
  noLabel?: boolean;
  /** Number format (w:numFmt, ST_NumberFormat) */
  numFmt?: (typeof NumberFormat)[keyof typeof NumberFormat];
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
  /** Default math font (m:mathFont) */
  mathFont?: string;
  /** Binary operator break style (m:brkBin) */
  binaryOperatorBreak?: "before" | "after" | "repeat";
  /** Subtraction binary operator break (m:brkBinSub) */
  binaryOperatorBreakSubtraction?: "--" | "-+" | "+-";
  /** Use small fractions (m:smallFrac) */
  smallFractions?: boolean;
  /** Use display defaults (m:dispDef) */
  displayDefaults?: boolean;
  /** Left margin in twips (m:lMargin) */
  leftMargin?: number;
  /** Right margin in twips (m:rMargin) */
  rightMargin?: number;
  /** Default justification (m:defJc) */
  defaultJustification?: "left" | "right" | "center" | "centerGroup";
  /** Spacing before a math instance in twips (m:preSp) */
  preSpacing?: number;
  /** Spacing after a math instance in twips (m:postSp) */
  postSpacing?: number;
  /** Inter-equation spacing in twips (m:interSp) */
  interSpacing?: number;
  /** Intra-equation spacing in twips (m:intraSp) */
  intraSpacing?: number;
  /** Wrap indent in twips (m:wrapIndent) */
  wrapIndent?: number;
  /** Wrap equations to the right (m:wrapRight; alternative to wrapIndent) */
  wrapRight?: boolean;
  /** Integral limit location (m:intLim) */
  integralLimitLocation?: "subSup" | "undOvr";
  /** N-ary limit location (m:naryLim) */
  naryLimitLocation?: "subSup" | "undOvr";
}
