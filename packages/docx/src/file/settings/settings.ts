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
  BuilderElement,
  XmlComponent,
  numberValObj,
  onOffObj,
  stringValObj,
} from "@file/xml-components";

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
  readonly compatibilityModeVersion?: number;
  /** Enable different headers/footers for even and odd pages */
  readonly evenAndOddHeaders?: boolean;
  /** Enable track changes (revision marking) */
  readonly trackRevisions?: boolean;
  /** Update fields when document is opened */
  readonly updateFields?: boolean;
  /** Compatibility settings for older Word versions */
  readonly compatibility?: CompatibilityOptions;
  /** Default distance between tab stops in twips */
  readonly defaultTabStop?: number;
  /** Hyphenation settings */
  readonly hyphenation?: HyphenationOptions;
  /** Controls whether punctuation is compressed at line ends */
  readonly characterSpacingControl?: "compressPunctuation" | "doNotCompress";
  /** Document protection settings */
  readonly documentProtection?: DocumentProtectionOptions;
  /** Default document view mode */
  readonly view?: "none" | "print" | "outline" | "masterPages" | "normal" | "web";
  /** Default zoom level (percentage) and type */
  readonly zoom?: {
    readonly percent?: number;
    readonly val?: "none" | "fullPage" | "bestFit" | "textFit";
  };
  /** Write protection recommendation (not enforcement) */
  readonly writeProtection?: WriteProtectionOptions;
  /** Whether to display the background shape in print layout */
  readonly displayBackgroundShape?: boolean;
  /** Whether to embed TrueType fonts in the document */
  readonly embedTrueTypeFonts?: boolean;
  /** Whether to embed system fonts in the document */
  readonly embedSystemFonts?: boolean;
  /** Whether to save only a subset of the embedded fonts */
  readonly saveSubsetFonts?: boolean;
  /** Document variables (key-value pairs stored in the document) */
  readonly docVars?: readonly { readonly name: string; readonly val: string }[];
  /** Mail merge configuration */
  readonly mailMerge?: MailMergeOptions;
  /** Theme color scheme remapping */
  readonly colorSchemeMapping?: {
    readonly bg1?: string;
    readonly t1?: string;
    readonly bg2?: string;
    readonly t2?: string;
    readonly accent1?: string;
    readonly accent2?: string;
    readonly accent3?: string;
    readonly accent4?: string;
    readonly accent5?: string;
    readonly accent6?: string;
    readonly hyperlink?: string;
    readonly followedHyperlink?: string;
  };
}

/**
 * Options for document protection (restrict editing).
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_DocProtect
 */
export interface DocumentProtectionOptions {
  /** Type of editing restriction */
  readonly edit?: "none" | "readOnly" | "comments" | "trackedChanges" | "forms";
  /** Whether formatting is restricted */
  readonly formatting?: boolean;
  /** Password hash (SHA-512 base64) */
  readonly hashValue?: string;
  /** Password salt (base64) */
  readonly saltValue?: string;
  /** Password spin count */
  readonly spinCount?: number;
  /** Password algorithm name */
  readonly algorithmName?: string;
}

/**
 * Options for write protection (read-only recommendation, not enforcement).
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_WriteProtection
 */
export interface WriteProtectionOptions {
  /** Cryptographic hash of the password */
  readonly hashValue?: string;
  /** Salt value for the hash (base64) */
  readonly saltValue?: string;
  /** Password spin count */
  readonly spinCount?: number;
  /** Password algorithm name */
  readonly algorithmName?: string;
  /** Whether write protection is recommended (default true when options provided) */
  readonly recommended?: boolean;
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
  readonly type?: OdsoFieldType;
  readonly name?: string;
  readonly mappedName?: string;
  readonly column?: number;
  readonly lid?: string;
  readonly dynamicAddress?: boolean;
}

/** Office Data Source Object (CT_Odso) */
export interface OdsoOptions {
  readonly udl?: string;
  readonly table?: string;
  readonly src?: string;
  readonly colDelim?: number;
  readonly type?: MailMergeSourceType;
  readonly fHdr?: boolean;
  readonly fieldMapData?: readonly OdsoFieldMapDataOptions[];
  readonly recipientData?: readonly string[];
}

/** Mail merge configuration (CT_MailMerge) */
export interface MailMergeOptions {
  /** Main document type (required) */
  readonly mainDocumentType: MailMergeDocType;
  /** Data source type (required) */
  readonly dataType: MailMergeDataType;
  /** Destination for merged documents */
  readonly destination?: MailMergeDest;
  /** Database connection string */
  readonly connectString?: string;
  /** SQL query to select data */
  readonly query?: string;
  /** Path to data source (relationship ID) */
  readonly dataSource?: string;
  /** Path to header source (relationship ID) */
  readonly headerSource?: string;
  /** Do not suppress blank lines */
  readonly doNotSuppressBlankLines?: boolean;
  /** Address field name for email merge */
  readonly addressFieldName?: string;
  /** Email subject line */
  readonly mailSubject?: string;
  /** Send as email attachment */
  readonly mailAsAttachment?: boolean;
  /** View merged data in document */
  readonly viewMergedData?: boolean;
  /** Active record index */
  readonly activeRecord?: number;
  /** Check errors mode */
  readonly checkErrors?: number;
  /** Link to query in data source */
  readonly linkToQuery?: boolean;
  /** Office Data Source Object configuration */
  readonly odso?: OdsoOptions;
}

/**
 * Options for automatic hyphenation settings.
 *
 * @see {@link Settings}
 */
export interface HyphenationOptions {
  /** Specifies whether the application automatically hyphenates words as they are typed in the document. */
  readonly autoHyphenation?: boolean;
  /** Specifies the minimum number of characters at the beginning of a word before a hyphen can be inserted. */
  readonly hyphenationZone?: number;
  /** Specifies the maximum number of consecutive lines that can end with a hyphenated word. */
  readonly consecutiveHyphenLimit?: number;
  /** Specifies whether to hyphenate words in all capital letters. */
  readonly doNotHyphenateCaps?: boolean;
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

    if (options.displayBackgroundShape !== undefined) {
      this.root.push(onOffObj("w:displayBackgroundShape", options.displayBackgroundShape));
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

    if (options.trackRevisions !== undefined) {
      this.root.push(onOffObj("w:trackRevisions", options.trackRevisions));
    }

    if (options.mailMerge !== undefined) {
      this.root.push(new MailMerge(options.mailMerge));
    }

    if (options.documentProtection !== undefined) {
      this.root.push(new DocumentProtection(options.documentProtection));
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

    if (options.evenAndOddHeaders !== undefined) {
      this.root.push(onOffObj("w:evenAndOddHeaders", options.evenAndOddHeaders));
    }

    this.root.push(
      stringValObj(
        "w:characterSpacingControl",
        options.characterSpacingControl ?? "compressPunctuation",
      ),
    );

    if (options.updateFields !== undefined) {
      this.root.push(onOffObj("w:updateFields", options.updateFields));
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

    if (options.colorSchemeMapping !== undefined) {
      this.root.push(new ColorSchemeMapping(options.colorSchemeMapping));
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
    if (options.algorithmName !== undefined) {
      attr["w:algorithmName"] = options.algorithmName;
    }
    if (options.hashValue !== undefined) {
      attr["w:hashValue"] = options.hashValue;
    }
    if (options.saltValue !== undefined) {
      attr["w:saltValue"] = options.saltValue;
    }
    if (options.spinCount !== undefined) {
      attr["w:spinCount"] = options.spinCount;
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
    if (options.hashValue !== undefined) {
      attr["w:hashValue"] = options.hashValue;
    }
    if (options.saltValue !== undefined) {
      attr["w:saltValue"] = options.saltValue;
    }
    if (options.spinCount !== undefined) {
      attr["w:spinCount"] = options.spinCount;
    }
    if (options.algorithmName !== undefined) {
      attr["w:algorithmName"] = options.algorithmName;
    }
    this.root.push({ _attr: attr });
  }
}

class ColorSchemeMapping extends XmlComponent {
  public constructor(options: {
    readonly bg1?: string;
    readonly t1?: string;
    readonly bg2?: string;
    readonly t2?: string;
    readonly accent1?: string;
    readonly accent2?: string;
    readonly accent3?: string;
    readonly accent4?: string;
    readonly accent5?: string;
    readonly accent6?: string;
    readonly hyperlink?: string;
    readonly followedHyperlink?: string;
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
