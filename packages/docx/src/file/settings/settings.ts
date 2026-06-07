import {
  FootnotePositionType,
  EndnotePositionType,
} from "@file/document/body/section-properties/properties/footnote-endnote-properties";
import {
  BuilderElement,
  XmlComponent,
  numberValObj,
  onOffObj,
  stringValObj,
} from "@file/xml-components";
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
import { derivePasswordHash } from "@office-open/core";

import { Compatibility } from "./compatibility";
import type { CompatibilityOptions } from "./compatibility";

// <xsd:complexType name="CT_Settings">
// <xsd:sequence>
//   <xsd:element name="writeProtection" type="CT_WriteProtection" minOccurs="0"/>
//   <xsd:element name="view" type="CT_View" minOccurs="0"/>
//   <xsd:element name="zoom" type="CT_Zoom" minOccurs="0"/>
//   <xsd:element name="removePersonalInformation" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="removeDateAndTime" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="doNotDisplayPageBoundaries" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="displayBackgroundShape" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="printPostScriptOverText" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="printFractionalCharacterWidth" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="printFormsData" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="embedTrueTypeFonts" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="embedSystemFonts" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="saveSubsetFonts" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="saveFormsData" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="mirrorMargins" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="alignBordersAndEdges" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="bordersDoNotSurroundHeader" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="bordersDoNotSurroundFooter" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="gutterAtTop" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="hideSpellingErrors" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="hideGrammaticalErrors" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="activeWritingStyle" type="CT_WritingStyle" minOccurs="0"
//     MaxOccurs="unbounded"/>
//   <xsd:element name="proofState" type="CT_Proof" minOccurs="0"/>
//   <xsd:element name="formsDesign" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="attachedTemplate" type="CT_Rel" minOccurs="0"/>
//   <xsd:element name="linkStyles" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="stylePaneFormatFilter" type="CT_StylePaneFilter" minOccurs="0"/>
//   <xsd:element name="stylePaneSortMethod" type="CT_StyleSort" minOccurs="0"/>
//   <xsd:element name="documentType" type="CT_DocType" minOccurs="0"/>
//   <xsd:element name="mailMerge" type="CT_MailMerge" minOccurs="0"/>
//   <xsd:element name="revisionView" type="CT_TrackChangesView" minOccurs="0"/>
//   <xsd:element name="trackRevisions" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="doNotTrackMoves" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="doNotTrackFormatting" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="documentProtection" type="CT_DocProtect" minOccurs="0"/>
//   <xsd:element name="autoFormatOverride" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="styleLockTheme" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="styleLockQFSet" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="defaultTabStop" type="CT_TwipsMeasure" minOccurs="0"/>
//   <xsd:element name="autoHyphenation" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="consecutiveHyphenLimit" type="CT_DecimalNumber" minOccurs="0"/>
//   <xsd:element name="hyphenationZone" type="CT_TwipsMeasure" minOccurs="0"/>
//   <xsd:element name="doNotHyphenateCaps" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="showEnvelope" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="summaryLength" type="CT_DecimalNumberOrPrecent" minOccurs="0"/>
//   <xsd:element name="clickAndTypeStyle" type="CT_String" minOccurs="0"/>
//   <xsd:element name="defaultTableStyle" type="CT_String" minOccurs="0"/>
//   <xsd:element name="evenAndOddHeaders" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="bookFoldRevPrinting" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="bookFoldPrinting" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="bookFoldPrintingSheets" type="CT_DecimalNumber" minOccurs="0"/>
//   <xsd:element name="drawingGridHorizontalSpacing" type="CT_TwipsMeasure" minOccurs="0"/>
//   <xsd:element name="drawingGridVerticalSpacing" type="CT_TwipsMeasure" minOccurs="0"/>
//   <xsd:element name="displayHorizontalDrawingGridEvery" type="CT_DecimalNumber" minOccurs="0"/>
//   <xsd:element name="displayVerticalDrawingGridEvery" type="CT_DecimalNumber" minOccurs="0"/>
//   <xsd:element name="doNotUseMarginsForDrawingGridOrigin" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="drawingGridHorizontalOrigin" type="CT_TwipsMeasure" minOccurs="0"/>
//   <xsd:element name="drawingGridVerticalOrigin" type="CT_TwipsMeasure" minOccurs="0"/>
//   <xsd:element name="doNotShadeFormData" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="noPunctuationKerning" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="characterSpacingControl" type="CT_CharacterSpacing" minOccurs="0"/>
//   <xsd:element name="printTwoOnOne" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="strictFirstAndLastChars" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="noLineBreaksAfter" type="CT_Kinsoku" minOccurs="0"/>
//   <xsd:element name="noLineBreaksBefore" type="CT_Kinsoku" minOccurs="0"/>
//   <xsd:element name="savePreviewPicture" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="doNotValidateAgainstSchema" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="saveInvalidXml" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="ignoreMixedContent" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="alwaysShowPlaceholderText" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="doNotDemarcateInvalidXml" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="saveXmlDataOnly" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="useXSLTWhenSaving" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="saveThroughXslt" type="CT_SaveThroughXslt" minOccurs="0"/>
//   <xsd:element name="showXMLTags" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="alwaysMergeEmptyNamespace" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="updateFields" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="hdrShapeDefaults" type="CT_ShapeDefaults" minOccurs="0"/>
//   <xsd:element name="footnotePr" type="CT_FtnDocProps" minOccurs="0"/>
//   <xsd:element name="endnotePr" type="CT_EdnDocProps" minOccurs="0"/>
//   <xsd:element name="compat" type="CT_Compat" minOccurs="0"/>
//   <xsd:element name="docVars" type="CT_DocVars" minOccurs="0"/>
//   <xsd:element name="rsids" type="CT_DocRsids" minOccurs="0"/>
//   <xsd:element ref="m:mathPr" minOccurs="0" maxOccurs="1"/>
//   <xsd:element name="attachedSchema" type="CT_String" minOccurs="0" maxOccurs="unbounded"/>
//   <xsd:element name="themeFontLang" type="CT_Language" minOccurs="0" maxOccurs="1"/>
//   <xsd:element name="clrSchemeMapping" type="CT_ColorSchemeMapping" minOccurs="0"/>
//   <xsd:element name="doNotIncludeSubdocsInStats" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="doNotAutoCompressPictures" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="forceUpgrade" type="CT_Empty" minOccurs="0" maxOccurs="1"/>
//   <xsd:element name="captions" type="CT_Captions" minOccurs="0" maxOccurs="1"/>
//   <xsd:element name="readModeInkLockDown" type="CT_ReadingModeInkLockDown" minOccurs="0"/>
//   <xsd:element name="smartTagType" type="CT_SmartTagType" minOccurs="0" maxOccurs="unbounded"/>
//   <xsd:element ref="sl:schemaLibrary" minOccurs="0" maxOccurs="1"/>
//   <xsd:element name="shapeDefaults" type="CT_ShapeDefaults" minOccurs="0"/>
//   <xsd:element name="doNotEmbedSmartTags" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="decimalSymbol" type="CT_String" minOccurs="0" maxOccurs="1"/>
//   <xsd:element name="listSeparator" type="CT_String" minOccurs="0" maxOccurs="1"/>
// </xsd:sequence>
// </xsd:complexType>

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

/**
 * Represents document settings in a WordprocessingML document.
 *
 * Settings contain document-wide configuration options such as
 * compatibility mode, track changes, hyphenation, and more.
 *
 * Reference: http://officeopenxml.com/WPsettings.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Settings">
 *   <xsd:sequence>
 *     <xsd:element name="trackRevisions" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="evenAndOddHeaders" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="defaultTabStop" type="CT_TwipsMeasure" minOccurs="0"/>
 *     <xsd:element name="autoHyphenation" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="consecutiveHyphenLimit" type="CT_DecimalNumber" minOccurs="0"/>
 *     <xsd:element name="hyphenationZone" type="CT_TwipsMeasure" minOccurs="0"/>
 *     <xsd:element name="doNotHyphenateCaps" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="updateFields" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="compat" type="CT_Compat" minOccurs="0"/>
 *     <!-- Additional elements omitted for brevity -->
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Basic settings with track changes enabled
 * new Settings({
 *   trackRevisions: true,
 *   evenAndOddHeaders: true,
 * });
 *
 * // Settings with compatibility mode and hyphenation
 * new Settings({
 *   compatibility: {
 *     version: 15, // Word 2013+
 *   },
 *   hyphenation: {
 *     autoHyphenation: true,
 *     consecutiveHyphenLimit: 2,
 *   },
 * });
 * ```
 */
export class Settings extends XmlComponent {
  public constructor(options: SettingsOptions) {
    super("w:settings");
    this.root.push({
      _attr: {
        "mc:Ignorable": "w14 w15 wp14",
        "xmlns:m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "xmlns:o": "urn:schemas-microsoft-com:office:office",
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xmlns:v": "urn:schemas-microsoft-com:vml",
        "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "xmlns:w10": "urn:schemas-microsoft-com:office:word",
        "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
        "xmlns:w15": "http://schemas.microsoft.com/office/word/2012/wordml",
        "xmlns:wne": "http://schemas.microsoft.com/office/word/2006/wordml",
        "xmlns:wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "xmlns:wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
        "xmlns:wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
        "xmlns:wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        "xmlns:wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
        "xmlns:wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
      },
    });

    // Elements ordered per XSD CT_Settings sequence.
    // Only elements actually used are included; all others are minOccurs="0".

    if (options.writeProtection !== undefined) {
      this.root.push(new WriteProtection(options.writeProtection));
    }

    if (options.view !== undefined) {
      this.root.push(new View(options.view));
    }

    if (options.zoom !== undefined) {
      this.root.push(
        new Zoom({
          val: options.zoom.val,
          percent: options.zoom.percent ?? 100,
        }),
      );
    }

    if (options.removePersonalInformation !== undefined) {
      this.root.push(onOffObj("w:removePersonalInformation", options.removePersonalInformation));
    }

    if (options.removeDateAndTime !== undefined) {
      this.root.push(onOffObj("w:removeDateAndTime", options.removeDateAndTime));
    }

    if (options.displayBackgroundShape !== undefined) {
      this.root.push(onOffObj("w:displayBackgroundShape", options.displayBackgroundShape));
    }

    if (options.doNotDisplayPageBoundaries !== undefined) {
      this.root.push(onOffObj("w:doNotDisplayPageBoundaries", options.doNotDisplayPageBoundaries));
    }

    if (options.printPostScriptOverText !== undefined) {
      this.root.push(onOffObj("w:printPostScriptOverText", options.printPostScriptOverText));
    }

    if (options.printFractionalCharacterWidth !== undefined) {
      this.root.push(
        onOffObj("w:printFractionalCharacterWidth", options.printFractionalCharacterWidth),
      );
    }

    if (options.printFormsData !== undefined) {
      this.root.push(onOffObj("w:printFormsData", options.printFormsData));
    }

    if (options.embedTrueTypeFonts !== undefined) {
      this.root.push(onOffObj("w:embedTrueTypeFonts", options.embedTrueTypeFonts));
    }

    if (options.embedSystemFonts !== undefined) {
      this.root.push(onOffObj("w:embedSystemFonts", options.embedSystemFonts));
    }

    if (options.saveSubsetFonts !== undefined) {
      this.root.push(onOffObj("w:saveSubsetFonts", options.saveSubsetFonts));
    }

    if (options.saveFormsData !== undefined) {
      this.root.push(onOffObj("w:saveFormsData", options.saveFormsData));
    }

    if (options.mirrorMargins !== undefined) {
      this.root.push(onOffObj("w:mirrorMargins", options.mirrorMargins));
    }

    if (options.alignBordersAndEdges !== undefined) {
      this.root.push(onOffObj("w:alignBordersAndEdges", options.alignBordersAndEdges));
    }

    if (options.bordersDoNotSurroundHeader !== undefined) {
      this.root.push(onOffObj("w:bordersDoNotSurroundHeader", options.bordersDoNotSurroundHeader));
    }

    if (options.bordersDoNotSurroundFooter !== undefined) {
      this.root.push(onOffObj("w:bordersDoNotSurroundFooter", options.bordersDoNotSurroundFooter));
    }

    if (options.gutterAtTop !== undefined) {
      this.root.push(onOffObj("w:gutterAtTop", options.gutterAtTop));
    }

    if (options.hideSpellingErrors !== undefined) {
      this.root.push(onOffObj("w:hideSpellingErrors", options.hideSpellingErrors));
    }

    if (options.hideGrammaticalErrors !== undefined) {
      this.root.push(onOffObj("w:hideGrammaticalErrors", options.hideGrammaticalErrors));
    }

    if (options.activeWritingStyle !== undefined) {
      for (const ws of options.activeWritingStyle) {
        const attrs: { key: string; value: string | boolean | number }[] = [];
        if (ws.lang !== undefined) attrs.push({ key: "w:lang", value: ws.lang });
        if (ws.vendorID !== undefined) attrs.push({ key: "w:vendorID", value: ws.vendorID });
        if (ws.dllVersion !== undefined) attrs.push({ key: "w:dllVersion", value: ws.dllVersion });
        if (ws.nlCheck !== undefined) attrs.push({ key: "w:nlCheck", value: ws.nlCheck });
        if (ws.checkStyle !== undefined) attrs.push({ key: "w:checkStyle", value: ws.checkStyle });
        if (ws.appCheck !== undefined) attrs.push({ key: "w:appCheck", value: ws.appCheck });
        if (ws.appName !== undefined) attrs.push({ key: "w:appName", value: ws.appName });
        this.root.push(new BuilderElement({ name: "w:activeWritingStyle", attributes: attrs }));
      }
    }

    if (options.proofState !== undefined) {
      const attrs: { key: string; value: string }[] = [];
      if (options.proofState.spelling !== undefined)
        attrs.push({ key: "w:spelling", value: options.proofState.spelling });
      if (options.proofState.grammar !== undefined)
        attrs.push({ key: "w:grammar", value: options.proofState.grammar });
      this.root.push(new BuilderElement({ name: "w:proofState", attributes: attrs }));
    }

    if (options.formsDesign !== undefined) {
      this.root.push(onOffObj("w:formsDesign", options.formsDesign));
    }

    if (options.attachedTemplate !== undefined) {
      this.root.push(
        new BuilderElement({
          name: "w:attachedTemplate",
          attributes: [{ key: "r:id", value: options.attachedTemplate }],
        }),
      );
    }

    if (options.linkStyles !== undefined) {
      this.root.push(onOffObj("w:linkStyles", options.linkStyles));
    }

    if (options.stylePaneFormatFilter !== undefined) {
      const f = options.stylePaneFormatFilter;
      const attrs: { key: string; value: string }[] = [];
      const boolFlags: { prop: keyof typeof f; xmlKey: string }[] = [
        { prop: "allStyles", xmlKey: "w:allStyles" },
        { prop: "customStyles", xmlKey: "w:customStyles" },
        { prop: "stylesInUse", xmlKey: "w:stylesInUse" },
        { prop: "headingStyles", xmlKey: "w:headingStyles" },
        { prop: "numberingStyles", xmlKey: "w:numberingStyles" },
        { prop: "tableStyles", xmlKey: "w:tableStyles" },
        { prop: "directFormattingOnRuns", xmlKey: "w:directFormattingOnRuns" },
        { prop: "directFormattingOnParagraphs", xmlKey: "w:directFormattingOnParagraphs" },
        { prop: "directFormattingOnNumbering", xmlKey: "w:directFormattingOnNumbering" },
        { prop: "directFormattingOnTables", xmlKey: "w:directFormattingOnTables" },
        { prop: "clearFormatting", xmlKey: "w:clearFormatting" },
        { prop: "top3HeadingStyles", xmlKey: "w:top3HeadingStyles" },
        { prop: "visibleStyles", xmlKey: "w:visibleStyles" },
        { prop: "alternateStyleNames", xmlKey: "w:alternateStyleNames" },
      ];
      for (const { prop, xmlKey } of boolFlags) {
        if (f[prop] !== undefined) attrs.push({ key: xmlKey, value: f[prop] ? "1" : "0" });
      }
      this.root.push(new BuilderElement({ name: "w:stylePaneFormatFilter", attributes: attrs }));
    }

    if (options.stylePaneSortMethod !== undefined) {
      this.root.push(stringValObj("w:stylePaneSortMethod", options.stylePaneSortMethod));
    }

    if (options.documentType !== undefined) {
      this.root.push(stringValObj("w:documentType", options.documentType));
    }

    if (options.revisionView !== undefined) {
      this.root.push(new RevisionView(options.revisionView));
    }

    if (options.trackRevisions !== undefined) {
      this.root.push(onOffObj("w:trackRevisions", options.trackRevisions));
    }

    if (options.doNotTrackMoves !== undefined) {
      this.root.push(onOffObj("w:doNotTrackMoves", options.doNotTrackMoves));
    }

    if (options.doNotTrackFormatting !== undefined) {
      this.root.push(onOffObj("w:doNotTrackFormatting", options.doNotTrackFormatting));
    }

    if (options.mailMerge !== undefined) {
      this.root.push(new MailMerge(options.mailMerge));
    }

    if (options.documentProtection !== undefined) {
      this.root.push(new DocumentProtection(options.documentProtection));
    }

    if (options.autoFormatOverride !== undefined) {
      this.root.push(onOffObj("w:autoFormatOverride", options.autoFormatOverride));
    }

    if (options.styleLockTheme !== undefined) {
      this.root.push(onOffObj("w:styleLockTheme", options.styleLockTheme));
    }

    if (options.styleLockQFSet !== undefined) {
      this.root.push(onOffObj("w:styleLockQFSet", options.styleLockQFSet));
    }

    // https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_defaultTabStop_topic_ID0EIXSX.html
    this.root.push(numberValObj("w:defaultTabStop", options.defaultTabStop ?? 420));

    if (options.hyphenation?.autoHyphenation !== undefined) {
      this.root.push(onOffObj("w:autoHyphenation", options.hyphenation.autoHyphenation));
    }

    if (options.hyphenation?.consecutiveHyphenLimit !== undefined) {
      this.root.push(
        numberValObj("w:consecutiveHyphenLimit", options.hyphenation.consecutiveHyphenLimit),
      );
    }

    if (options.hyphenation?.hyphenationZone !== undefined) {
      this.root.push(numberValObj("w:hyphenationZone", options.hyphenation.hyphenationZone));
    }

    if (options.hyphenation?.doNotHyphenateCaps !== undefined) {
      this.root.push(onOffObj("w:doNotHyphenateCaps", options.hyphenation.doNotHyphenateCaps));
    }

    if (options.showEnvelope !== undefined) {
      this.root.push(onOffObj("w:showEnvelope", options.showEnvelope));
    }

    if (options.summaryLength !== undefined) {
      this.root.push(numberValObj("w:summaryLength", options.summaryLength));
    }

    if (options.clickAndTypeStyle !== undefined) {
      this.root.push(stringValObj("w:clickAndTypeStyle", options.clickAndTypeStyle));
    }

    if (options.defaultTableStyle !== undefined) {
      this.root.push(stringValObj("w:defaultTableStyle", options.defaultTableStyle));
    }

    if (options.evenAndOddHeaders !== undefined) {
      this.root.push(onOffObj("w:evenAndOddHeaders", options.evenAndOddHeaders));
    }

    if (options.bookFoldRevPrinting !== undefined) {
      this.root.push(onOffObj("w:bookFoldRevPrinting", options.bookFoldRevPrinting));
    }

    if (options.bookFoldPrinting !== undefined) {
      this.root.push(onOffObj("w:bookFoldPrinting", options.bookFoldPrinting));
    }

    if (options.bookFoldPrintingSheets !== undefined) {
      this.root.push(numberValObj("w:bookFoldPrintingSheets", options.bookFoldPrintingSheets));
    }

    if (options.drawingGridHorizontalSpacing !== undefined) {
      this.root.push(
        numberValObj("w:drawingGridHorizontalSpacing", options.drawingGridHorizontalSpacing),
      );
    }

    if (options.drawingGridVerticalSpacing !== undefined) {
      this.root.push(
        numberValObj("w:drawingGridVerticalSpacing", options.drawingGridVerticalSpacing),
      );
    }

    if (options.displayHorizontalDrawingGridEvery !== undefined) {
      this.root.push(
        numberValObj(
          "w:displayHorizontalDrawingGridEvery",
          options.displayHorizontalDrawingGridEvery,
        ),
      );
    }

    if (options.displayVerticalDrawingGridEvery !== undefined) {
      this.root.push(
        numberValObj("w:displayVerticalDrawingGridEvery", options.displayVerticalDrawingGridEvery),
      );
    }

    if (options.drawingGridHorizontalOrigin !== undefined) {
      this.root.push(
        numberValObj("w:drawingGridHorizontalOrigin", options.drawingGridHorizontalOrigin),
      );
    }

    if (options.drawingGridVerticalOrigin !== undefined) {
      this.root.push(
        numberValObj("w:drawingGridVerticalOrigin", options.drawingGridVerticalOrigin),
      );
    }

    if (options.doNotUseMarginsForDrawingGridOrigin !== undefined) {
      this.root.push(
        onOffObj(
          "w:doNotUseMarginsForDrawingGridOrigin",
          options.doNotUseMarginsForDrawingGridOrigin,
        ),
      );
    }

    if (options.doNotShadeFormData !== undefined) {
      this.root.push(onOffObj("w:doNotShadeFormData", options.doNotShadeFormData));
    }

    this.root.push(
      stringValObj(
        "w:characterSpacingControl",
        options.characterSpacingControl ?? "compressPunctuation",
      ),
    );

    if (options.noPunctuationKerning !== undefined) {
      this.root.push(onOffObj("w:noPunctuationKerning", options.noPunctuationKerning));
    }

    if (options.printTwoOnOne !== undefined) {
      this.root.push(onOffObj("w:printTwoOnOne", options.printTwoOnOne));
    }

    if (options.strictFirstAndLastChars !== undefined) {
      this.root.push(onOffObj("w:strictFirstAndLastChars", options.strictFirstAndLastChars));
    }

    if (options.noLineBreaksAfter !== undefined) {
      const attrs: { key: string; value: string }[] = [];
      if (options.noLineBreaksAfter.lang !== undefined)
        attrs.push({ key: "w:lang", value: options.noLineBreaksAfter.lang });
      if (options.noLineBreaksAfter.val !== undefined)
        attrs.push({ key: "w:val", value: options.noLineBreaksAfter.val });
      this.root.push(new BuilderElement({ name: "w:noLineBreaksAfter", attributes: attrs }));
    }

    if (options.noLineBreaksBefore !== undefined) {
      const attrs: { key: string; value: string }[] = [];
      if (options.noLineBreaksBefore.lang !== undefined)
        attrs.push({ key: "w:lang", value: options.noLineBreaksBefore.lang });
      if (options.noLineBreaksBefore.val !== undefined)
        attrs.push({ key: "w:val", value: options.noLineBreaksBefore.val });
      this.root.push(new BuilderElement({ name: "w:noLineBreaksBefore", attributes: attrs }));
    }

    if (options.savePreviewPicture !== undefined) {
      this.root.push(onOffObj("w:savePreviewPicture", options.savePreviewPicture));
    }

    if (options.doNotValidateAgainstSchema !== undefined) {
      this.root.push(onOffObj("w:doNotValidateAgainstSchema", options.doNotValidateAgainstSchema));
    }

    if (options.saveInvalidXml !== undefined) {
      this.root.push(onOffObj("w:saveInvalidXml", options.saveInvalidXml));
    }

    if (options.ignoreMixedContent !== undefined) {
      this.root.push(onOffObj("w:ignoreMixedContent", options.ignoreMixedContent));
    }

    if (options.alwaysShowPlaceholderText !== undefined) {
      this.root.push(onOffObj("w:alwaysShowPlaceholderText", options.alwaysShowPlaceholderText));
    }

    if (options.doNotDemarcateInvalidXml !== undefined) {
      this.root.push(onOffObj("w:doNotDemarcateInvalidXml", options.doNotDemarcateInvalidXml));
    }

    if (options.saveXmlDataOnly !== undefined) {
      this.root.push(onOffObj("w:saveXmlDataOnly", options.saveXmlDataOnly));
    }

    if (options.useXSLTWhenSaving !== undefined) {
      this.root.push(onOffObj("w:useXSLTWhenSaving", options.useXSLTWhenSaving));
    }

    if (options.saveThroughXslt !== undefined) {
      const attrs: { key: string; value: string }[] = [];
      if (options.saveThroughXslt.id !== undefined)
        attrs.push({ key: "r:id", value: options.saveThroughXslt.id });
      if (options.saveThroughXslt.val !== undefined)
        attrs.push({ key: "w:val", value: options.saveThroughXslt.val });
      if (options.saveThroughXslt.solutionID !== undefined)
        attrs.push({ key: "w:solutionID", value: options.saveThroughXslt.solutionID });
      this.root.push(new BuilderElement({ name: "w:saveThroughXslt", attributes: attrs }));
    }

    if (options.showXMLTags !== undefined) {
      this.root.push(onOffObj("w:showXMLTags", options.showXMLTags));
    }

    if (options.alwaysMergeEmptyNamespace !== undefined) {
      this.root.push(onOffObj("w:alwaysMergeEmptyNamespace", options.alwaysMergeEmptyNamespace));
    }

    if (options.updateFields !== undefined) {
      this.root.push(onOffObj("w:updateFields", options.updateFields));
    }

    if (options.hdrShapeDefaults !== undefined) {
      this.root.push(new BuilderElement({ name: "w:hdrShapeDefaults" }));
    }

    if (options.footnotePr !== undefined) {
      this.root.push(new FootnotePrElement(options.footnotePr));
    }

    if (options.endnotePr !== undefined) {
      this.root.push(new EndnotePrElement(options.endnotePr));
    }

    this.root.push(
      new Compatibility({
        ...options.compatibility,
        version: options.compatibility?.version ?? options.compatibilityModeVersion ?? 15,
        spaceForUnderline: options.compatibility?.spaceForUnderline ?? true,
        balanceSingleByteDoubleByteWidth:
          options.compatibility?.balanceSingleByteDoubleByteWidth ?? true,
        doNotLeaveBackslashAlone: options.compatibility?.doNotLeaveBackslashAlone ?? true,
        underlineTrailingSpaces: options.compatibility?.underlineTrailingSpaces ?? true,
        doNotExpandShiftReturn: options.compatibility?.doNotExpandShiftReturn ?? true,
        adjustLineHeightInTable: options.compatibility?.adjustLineHeightInTable ?? true,
        useFELayout: options.compatibility?.useFELayout ?? true,
        overrideTableStyleFontSizeAndJustification:
          options.compatibility?.overrideTableStyleFontSizeAndJustification ?? true,
        enableOpenTypeFeatures: options.compatibility?.enableOpenTypeFeatures ?? true,
        doNotFlipMirrorIndents: options.compatibility?.doNotFlipMirrorIndents ?? true,
      }),
    );

    if (options.docVars !== undefined && options.docVars.length > 0) {
      this.root.push(
        new BuilderElement({
          name: "w:docVars",
          children: options.docVars.map(
            (v) =>
              new BuilderElement({
                name: "w:docVar",
                attributes: [
                  { key: "w:name", value: v.name },
                  { key: "w:val", value: v.val },
                ],
              }),
          ),
        }),
      );
    }

    if (options.rsids !== undefined) {
      this.root.push(new RsidsElement(options.rsids));
    }

    if (options.mathPr !== undefined) {
      this.root.push(new MathPrElement(options.mathPr));
    }

    if (options.attachedSchema !== undefined) {
      for (const schema of options.attachedSchema) {
        this.root.push(stringValObj("w:attachedSchema", schema));
      }
    }

    if (options.colorSchemeMapping !== undefined) {
      this.root.push(new ColorSchemeMapping(options.colorSchemeMapping));
    }

    if (options.themeFontLang !== undefined) {
      this.root.push(
        new BuilderElement({
          name: "w:themeFontLang",
          attributes: [{ key: "w:val", value: options.themeFontLang }],
        }),
      );
    }

    if (options.doNotIncludeSubdocsInStats !== undefined) {
      this.root.push(onOffObj("w:doNotIncludeSubdocsInStats", options.doNotIncludeSubdocsInStats));
    }

    if (options.doNotAutoCompressPictures !== undefined) {
      this.root.push(onOffObj("w:doNotAutoCompressPictures", options.doNotAutoCompressPictures));
    }

    if (options.forceUpgrade !== undefined) {
      this.root.push(new BuilderElement({ name: "w:forceUpgrade" }));
    }

    if (options.captions !== undefined) {
      this.root.push(new CaptionsElement(options.captions));
    }

    if (options.readModeInkLockDown !== undefined) {
      this.root.push(new ReadModeInkLockDownElement(options.readModeInkLockDown));
    }

    if (options.smartTagType !== undefined) {
      for (const st of options.smartTagType) {
        const attrs: { key: string; value: string }[] = [];
        if (st.namespace !== undefined) attrs.push({ key: "w:namespace", value: st.namespace });
        if (st.namespaceuri !== undefined)
          attrs.push({ key: "w:namespaceuri", value: st.namespaceuri });
        if (st.name !== undefined) attrs.push({ key: "w:name", value: st.name });
        if (st.url !== undefined) attrs.push({ key: "w:url", value: st.url });
        this.root.push(new BuilderElement({ name: "w:smartTagType", attributes: attrs }));
      }
    }

    if (options.doNotEmbedSmartTags !== undefined) {
      this.root.push(onOffObj("w:doNotEmbedSmartTags", options.doNotEmbedSmartTags));
    }

    if (options.shapeDefaults !== undefined) {
      this.root.push(new BuilderElement({ name: "w:shapeDefaults" }));
    }

    if (options.decimalSymbol !== undefined) {
      this.root.push(stringValObj("w:decimalSymbol", options.decimalSymbol));
    }

    if (options.listSeparator !== undefined) {
      this.root.push(stringValObj("w:listSeparator", options.listSeparator));
    }
  }
}

/**
 * Represents document protection settings (CT_DocProtect).
 *
 * Restricts the types of editing allowed in the document.
 * Requires `enforcement: true` to take effect.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_DocProtect
 */
class DocumentProtection extends XmlComponent {
  public constructor(options: DocumentProtectionOptions) {
    super("w:documentProtection");
    const attr: Record<string, string | number> = { "w:enforcement": "1" };
    if (options.edit !== undefined) {
      attr["w:edit"] = options.edit;
    }
    if (options.formatting !== undefined) {
      attr["w:formatting"] = options.formatting ? "1" : "0";
    }
    // Auto-derive hash from plaintext password
    let derived: ReturnType<typeof derivePasswordHash> | undefined;
    if (options.password !== undefined && options.hashValue === undefined) {
      derived = derivePasswordHash(options.password);
    }
    const hashValue = options.hashValue ?? derived?.hashValue;
    const saltValue = options.saltValue ?? derived?.saltValue;
    const spinCount = options.spinCount ?? derived?.spinCount;
    const algorithmName = options.algorithmName ?? derived?.algorithmName;
    if (algorithmName !== undefined) {
      attr["w:algorithmName"] = algorithmName;
    }
    if (hashValue !== undefined) {
      attr["w:hashValue"] = hashValue;
    }
    if (saltValue !== undefined) {
      attr["w:saltValue"] = saltValue;
    }
    if (options.hash !== undefined) {
      attr["w:hash"] = options.hash;
    }
    if (options.salt !== undefined) {
      attr["w:salt"] = options.salt;
    }
    if (spinCount !== undefined) {
      attr["w:spinCount"] = spinCount;
    }
    if (options.cryptoAlgorithmClass !== undefined) {
      attr["w:cryptAlgorithmClass"] = options.cryptoAlgorithmClass;
    }
    if (options.cryptoAlgorithmSid !== undefined) {
      attr["w:cryptAlgorithmSid"] = options.cryptoAlgorithmSid;
    }
    if (options.cryptoAlgorithmType !== undefined) {
      attr["w:cryptAlgorithmType"] = options.cryptoAlgorithmType;
    }
    if (options.cryptoProvider !== undefined) {
      attr["w:cryptProvider"] = options.cryptoProvider;
    }
    if (options.cryptoProviderType !== undefined) {
      attr["w:cryptProviderType"] = options.cryptoProviderType;
    }
    if (options.cryptoProviderTypeExtension !== undefined) {
      attr["w:cryptProviderTypeExt"] = options.cryptoProviderTypeExtension;
    }
    if (options.cryptoProviderTypeExtensionSource !== undefined) {
      attr["w:cryptProviderTypeExtSource"] = options.cryptoProviderTypeExtensionSource;
    }
    if (options.algorithmExtensionId !== undefined) {
      attr["w:algIdExt"] = options.algorithmExtensionId;
    }
    if (options.algorithmExtensionSource !== undefined) {
      attr["w:algIdExtSource"] = options.algorithmExtensionSource;
    }
    if (options.cryptoSpinCount !== undefined) {
      attr["w:cryptSpinCount"] = options.cryptoSpinCount;
    }
    this.root.push({ _attr: attr });
  }
}

class View extends XmlComponent {
  public constructor(val: string) {
    super("w:view");
    this.root.push({ _attr: { "w:val": val } });
  }
}

class Zoom extends XmlComponent {
  public constructor(options: { val?: string; percent: number }) {
    super("w:zoom");
    const attr: Record<string, string | number> = { "w:percent": options.percent };
    if (options.val !== undefined) {
      attr["w:val"] = options.val;
    }
    this.root.push({ _attr: attr });
  }
}

class WriteProtection extends XmlComponent {
  public constructor(options: WriteProtectionOptions) {
    super("w:writeProtection");
    const attr: Record<string, string | number> = {};
    if (options.recommended !== undefined) {
      attr["w:recommended"] = options.recommended ? "1" : "0";
    }
    // Auto-derive hash from plaintext password
    let derived: ReturnType<typeof derivePasswordHash> | undefined;
    if (options.password !== undefined && options.hashValue === undefined) {
      derived = derivePasswordHash(options.password);
    }
    const hashValue = options.hashValue ?? derived?.hashValue;
    const saltValue = options.saltValue ?? derived?.saltValue;
    const spinCount = options.spinCount ?? derived?.spinCount;
    const algorithmName = options.algorithmName ?? derived?.algorithmName;
    if (hashValue !== undefined) {
      attr["w:hashValue"] = hashValue;
    }
    if (saltValue !== undefined) {
      attr["w:saltValue"] = saltValue;
    }
    if (options.hash !== undefined) {
      attr["w:hash"] = options.hash;
    }
    if (options.salt !== undefined) {
      attr["w:salt"] = options.salt;
    }
    if (spinCount !== undefined) {
      attr["w:spinCount"] = spinCount;
    }
    if (algorithmName !== undefined) {
      attr["w:algorithmName"] = algorithmName;
    }
    if (options.cryptoAlgorithmClass !== undefined) {
      attr["w:cryptAlgorithmClass"] = options.cryptoAlgorithmClass;
    }
    if (options.cryptoAlgorithmSid !== undefined) {
      attr["w:cryptAlgorithmSid"] = options.cryptoAlgorithmSid;
    }
    if (options.cryptoAlgorithmType !== undefined) {
      attr["w:cryptAlgorithmType"] = options.cryptoAlgorithmType;
    }
    if (options.cryptoProvider !== undefined) {
      attr["w:cryptProvider"] = options.cryptoProvider;
    }
    if (options.cryptoProviderType !== undefined) {
      attr["w:cryptProviderType"] = options.cryptoProviderType;
    }
    if (options.algorithmExtensionId !== undefined) {
      attr["w:algIdExt"] = options.algorithmExtensionId;
    }
    if (options.algorithmExtensionSource !== undefined) {
      attr["w:algIdExtSource"] = options.algorithmExtensionSource;
    }
    if (options.cryptoProviderTypeExtension !== undefined) {
      attr["w:cryptProviderTypeExt"] = options.cryptoProviderTypeExtension;
    }
    if (options.cryptoProviderTypeExtensionSource !== undefined) {
      attr["w:cryptProviderTypeExtSource"] = options.cryptoProviderTypeExtensionSource;
    }
    if (options.cryptoSpinCount !== undefined) {
      attr["w:cryptSpinCount"] = options.cryptoSpinCount;
    }
    this.root.push({ _attr: attr });
  }
}

class ColorSchemeMapping extends XmlComponent {
  public constructor(options: {
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
  }) {
    super("w:clrSchemeMapping");
    const attr: Record<string, string> = {};
    if (options.bg1 !== undefined) attr["w:bg1"] = options.bg1;
    if (options.t1 !== undefined) attr["w:t1"] = options.t1;
    if (options.bg2 !== undefined) attr["w:bg2"] = options.bg2;
    if (options.t2 !== undefined) attr["w:t2"] = options.t2;
    if (options.accent1 !== undefined) attr["w:accent1"] = options.accent1;
    if (options.accent2 !== undefined) attr["w:accent2"] = options.accent2;
    if (options.accent3 !== undefined) attr["w:accent3"] = options.accent3;
    if (options.accent4 !== undefined) attr["w:accent4"] = options.accent4;
    if (options.accent5 !== undefined) attr["w:accent5"] = options.accent5;
    if (options.accent6 !== undefined) attr["w:accent6"] = options.accent6;
    if (options.hyperlink !== undefined) attr["w:hyperlink"] = options.hyperlink;
    if (options.followedHyperlink !== undefined)
      attr["w:followedHyperlink"] = options.followedHyperlink;
    this.root.push({ _attr: attr });
  }
}

/**
 * Revision view settings (CT_TrackChangesView).
 * Controls which types of revisions are visible in the document.
 */
class RevisionView extends XmlComponent {
  public constructor(options: RevisionViewOptions) {
    super("w:revisionView");
    const attr: Record<string, string> = {};
    if (options.markup !== undefined) attr["w:markup"] = options.markup ? "true" : "false";
    if (options.comments !== undefined) attr["w:comments"] = options.comments ? "true" : "false";
    if (options.insDel !== undefined) attr["w:insDel"] = options.insDel ? "true" : "false";
    if (options.formatting !== undefined)
      attr["w:formatting"] = options.formatting ? "true" : "false";
    if (options.inkAnnotations !== undefined)
      attr["w:inkAnnotations"] = options.inkAnnotations ? "true" : "false";
    this.root.push({ _attr: attr });
  }
}

/**
 * Mail merge configuration (CT_MailMerge).
 *
 * Generates the w:mailMerge element in document settings with
 * mainDocumentType, dataType, connectString, query, ODSO, etc.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_MailMerge
 */
class MailMerge extends XmlComponent {
  public constructor(options: MailMergeOptions) {
    super("w:mailMerge");

    // XSD sequence order: mainDocumentType, linkToQuery, dataType, connectString,
    // query, dataSource, headerSource, doNotSuppressBlankLines, destination,
    // addressFieldName, mailSubject, mailAsAttachment, viewMergedData,
    // activeRecord, checkErrors, odso
    this.root.push(stringValObj("w:mainDocumentType", options.mainDocumentType));

    if (options.linkToQuery !== undefined) {
      this.root.push(onOffObj("w:linkToQuery", options.linkToQuery));
    }

    this.root.push(stringValObj("w:dataType", options.dataType));

    if (options.connectString !== undefined) {
      this.root.push(stringValObj("w:connectString", options.connectString));
    }

    if (options.query !== undefined) {
      this.root.push(stringValObj("w:query", options.query));
    }

    if (options.dataSource !== undefined) {
      this.root.push(
        new BuilderElement({
          name: "w:dataSource",
          attributes: [{ key: "r:id", value: options.dataSource }],
        }),
      );
    }

    if (options.headerSource !== undefined) {
      this.root.push(
        new BuilderElement({
          name: "w:headerSource",
          attributes: [{ key: "r:id", value: options.headerSource }],
        }),
      );
    }

    if (options.doNotSuppressBlankLines !== undefined) {
      this.root.push(onOffObj("w:doNotSuppressBlankLines", options.doNotSuppressBlankLines));
    }

    if (options.destination !== undefined) {
      this.root.push(stringValObj("w:destination", options.destination));
    }

    if (options.addressFieldName !== undefined) {
      this.root.push(stringValObj("w:addressFieldName", options.addressFieldName));
    }

    if (options.mailSubject !== undefined) {
      this.root.push(stringValObj("w:mailSubject", options.mailSubject));
    }

    if (options.mailAsAttachment !== undefined) {
      this.root.push(onOffObj("w:mailAsAttachment", options.mailAsAttachment));
    }

    if (options.viewMergedData !== undefined) {
      this.root.push(onOffObj("w:viewMergedData", options.viewMergedData));
    }

    if (options.activeRecord !== undefined) {
      this.root.push(numberValObj("w:activeRecord", options.activeRecord));
    }

    if (options.checkErrors !== undefined) {
      this.root.push(numberValObj("w:checkErrors", options.checkErrors));
    }

    if (options.odso !== undefined) {
      this.root.push(new Odso(options.odso));
    }

    if (options.active !== undefined) {
      this.root.push(onOffObj("w:active", options.active));
    }

    if (options.recipients !== undefined) {
      this.root.push(
        new BuilderElement({
          name: "w:recipients",
          attributes: [{ key: "r:id", value: options.recipients }],
        }),
      );
    }
  }
}

/**
 * Office Data Source Object (CT_Odso).
 */
class Odso extends XmlComponent {
  public constructor(options: OdsoOptions) {
    super("w:odso");

    if (options.udl !== undefined) {
      this.root.push(stringValObj("w:udl", options.udl));
    }
    if (options.table !== undefined) {
      this.root.push(stringValObj("w:table", options.table));
    }
    if (options.src !== undefined) {
      this.root.push(
        new BuilderElement({
          name: "w:src",
          attributes: [{ key: "r:id", value: options.src }],
        }),
      );
    }
    if (options.colDelim !== undefined) {
      this.root.push(numberValObj("w:colDelim", options.colDelim));
    }
    if (options.type !== undefined) {
      this.root.push(stringValObj("w:type", options.type));
    }
    if (options.fHdr !== undefined) {
      this.root.push(onOffObj("w:fHdr", options.fHdr));
    }

    if (options.fieldMapData !== undefined) {
      for (const fm of options.fieldMapData) {
        this.root.push(new OdsoFieldMapData(fm));
      }
    }

    if (options.recipientData !== undefined) {
      for (const rd of options.recipientData) {
        this.root.push(
          new BuilderElement({
            name: "w:recipientData",
            attributes: [{ key: "r:id", value: rd }],
          }),
        );
      }
    }

    if (options.uniqueTag !== undefined) {
      this.root.push(stringValObj("w:uniqueTag", options.uniqueTag));
    }
  }
}

/**
 * ODSO field map data (CT_OdsoFieldMapData).
 */
class OdsoFieldMapData extends XmlComponent {
  public constructor(options: OdsoFieldMapDataOptions) {
    super("w:fieldMapData");

    if (options.type !== undefined) {
      this.root.push(stringValObj("w:type", options.type));
    }
    if (options.name !== undefined) {
      this.root.push(stringValObj("w:name", options.name));
    }
    if (options.mappedName !== undefined) {
      this.root.push(stringValObj("w:mappedName", options.mappedName));
    }
    if (options.column !== undefined) {
      this.root.push(numberValObj("w:column", options.column));
    }
    if (options.lid !== undefined) {
      this.root.push(stringValObj("w:lid", options.lid));
    }
    if (options.dynamicAddress !== undefined) {
      this.root.push(onOffObj("w:dynamicAddress", options.dynamicAddress));
    }
  }
}

// ── Batch C: Complex substructure elements ──

/** Document-level footnote properties (CT_FtnDocProps) */
class FootnotePrElement extends XmlComponent {
  public constructor(options: FootnotePropertiesOptions) {
    super("w:footnotePr");

    if (options.pos !== undefined) {
      this.root.push(
        new BuilderElement({
          name: "w:pos",
          attributes: [{ key: "w:val", value: options.pos }],
        }),
      );
    }

    if (options.numFmt !== undefined || options.format !== undefined) {
      const attrs: { key: string; value: string }[] = [];
      if (options.numFmt !== undefined) attrs.push({ key: "w:val", value: options.numFmt });
      if (options.format !== undefined) attrs.push({ key: "w:format", value: options.format });
      this.root.push(new BuilderElement({ name: "w:numFmt", attributes: attrs }));
    }

    if (options.numStart !== undefined) {
      this.root.push(numberValObj("w:numStart", options.numStart));
    }

    if (options.numRestart !== undefined) {
      this.root.push(stringValObj("w:numRestart", options.numRestart));
    }
  }
}

/** Document-level endnote properties (CT_EdnDocProps) */
class EndnotePrElement extends XmlComponent {
  public constructor(options: EndnotePropertiesOptions) {
    super("w:endnotePr");

    if (options.pos !== undefined) {
      this.root.push(
        new BuilderElement({
          name: "w:pos",
          attributes: [{ key: "w:val", value: options.pos }],
        }),
      );
    }

    if (options.numFmt !== undefined || options.format !== undefined) {
      const attrs: { key: string; value: string }[] = [];
      if (options.numFmt !== undefined) attrs.push({ key: "w:val", value: options.numFmt });
      if (options.format !== undefined) attrs.push({ key: "w:format", value: options.format });
      this.root.push(new BuilderElement({ name: "w:numFmt", attributes: attrs }));
    }

    if (options.numStart !== undefined) {
      this.root.push(numberValObj("w:numStart", options.numStart));
    }

    if (options.numRestart !== undefined) {
      this.root.push(stringValObj("w:numRestart", options.numRestart));
    }
  }
}

/** Document revision save IDs (CT_DocRsids) */
class RsidsElement extends XmlComponent {
  public constructor(options: RsidsOptions) {
    super("w:rsids");

    if (options.rsidRoot !== undefined) {
      this.root.push(
        new BuilderElement({
          name: "w:rsidRoot",
          attributes: [{ key: "w:val", value: options.rsidRoot }],
        }),
      );
    }

    if (options.rsids !== undefined) {
      for (const rsid of options.rsids) {
        this.root.push(
          new BuilderElement({
            name: "w:rsid",
            attributes: [{ key: "w:val", value: rsid }],
          }),
        );
      }
    }
  }
}

/** Reading mode ink lock-down (CT_ReadingModeInkLockDown) */
class ReadModeInkLockDownElement extends XmlComponent {
  public constructor(options: ReadModeInkLockDownOptions) {
    super("w:readModeInkLockDown");
    const attrs: Record<string, string | number> = {
      "w:actualPg": options.actualPg === false ? "0" : "1",
      "w:w": options.w,
      "w:h": options.h,
      "w:fontSz": options.fontSz,
    };
    this.root.push({ _attr: attrs });
  }
}

/** Captions (CT_Captions) */
class CaptionsElement extends XmlComponent {
  public constructor(options: CaptionsOptions) {
    super("w:captions");

    for (const cap of options.captions) {
      const attrs: { key: string; value: string | number }[] = [{ key: "w:name", value: cap.name }];
      if (cap.pos !== undefined) attrs.push({ key: "w:pos", value: cap.pos });
      if (cap.chapNum !== undefined)
        attrs.push({ key: "w:chapNum", value: cap.chapNum ? "1" : "0" });
      if (cap.heading !== undefined) attrs.push({ key: "w:heading", value: cap.heading });
      if (cap.noLabel !== undefined)
        attrs.push({ key: "w:noLabel", value: cap.noLabel ? "1" : "0" });
      if (cap.numFmt !== undefined) attrs.push({ key: "w:numFmt", value: cap.numFmt });
      if (cap.sep !== undefined) attrs.push({ key: "w:sep", value: cap.sep });
      this.root.push(new BuilderElement({ name: "w:caption", attributes: attrs }));
    }

    if (options.autoCaptions !== undefined && options.autoCaptions.length > 0) {
      const autoCapChildren = options.autoCaptions.map(
        (ac) =>
          new BuilderElement({
            name: "w:autoCaption",
            attributes: [
              { key: "w:name", value: ac.name },
              { key: "w:caption", value: ac.caption },
            ],
          }),
      );
      this.root.push(new BuilderElement({ name: "w:autoCaptions", children: autoCapChildren }));
    }
  }
}

/** Math properties (m:mathPr) — uses math namespace */
class MathPrElement extends XmlComponent {
  public constructor(options: MathPropertiesOptions) {
    super("m:mathPr");

    if (options.mathFont !== undefined) {
      this.root.push(
        new BuilderElement({
          name: "m:mathFont",
          attributes: [{ key: "m:val", value: options.mathFont }],
        }),
      );
    }

    if (options.brkBin !== undefined) {
      this.root.push(
        new BuilderElement({
          name: "m:brkBin",
          attributes: [{ key: "m:val", value: options.brkBin }],
        }),
      );
    }

    if (options.brkBinSub !== undefined) {
      this.root.push(
        new BuilderElement({
          name: "m:brkBinSub",
          attributes: [{ key: "m:val", value: options.brkBinSub }],
        }),
      );
    }

    if (options.smallFrac !== undefined) {
      this.root.push(onOffObj("m:smallFrac", options.smallFrac));
    }

    if (options.dispDef !== undefined) {
      this.root.push(onOffObj("m:dispDef", options.dispDef));
    }

    if (options.lMargin !== undefined) {
      this.root.push(numberValObj("m:lMargin", options.lMargin));
    }

    if (options.rMargin !== undefined) {
      this.root.push(numberValObj("m:rMargin", options.rMargin));
    }

    if (options.defJc !== undefined) {
      this.root.push(
        new BuilderElement({
          name: "m:defJc",
          attributes: [{ key: "m:val", value: options.defJc }],
        }),
      );
    }

    if (options.wrapIndent !== undefined) {
      this.root.push(numberValObj("m:wrapIndent", options.wrapIndent));
    }

    if (options.intLim !== undefined) {
      this.root.push(
        new BuilderElement({
          name: "m:intLim",
          attributes: [{ key: "m:val", value: options.intLim }],
        }),
      );
    }

    if (options.naryLim !== undefined) {
      this.root.push(
        new BuilderElement({
          name: "m:naryLim",
          attributes: [{ key: "m:val", value: options.naryLim }],
        }),
      );
    }
  }
}
