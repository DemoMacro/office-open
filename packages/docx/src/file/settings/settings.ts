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
    NumberValueElement,
    OnOffElement,
    StringValueElement,
    XmlAttributeComponent,
    XmlComponent,
} from "@file/xml-components";

import { Compatibility } from "./compatibility";
import type { ICompatibilityOptions } from "./compatibility";

/**
 * Attributes for the settings element with XML namespace declarations.
 *
 * Defines the XML namespaces required for the settings.xml document part.
 * These namespaces enable compatibility features, markup compatibility,
 * and various Office-specific extensions.
 *
 * @internal
 */
export class SettingsAttributes extends XmlAttributeComponent<{
    readonly wpc?: string;
    readonly mc?: string;
    readonly o?: string;
    readonly r?: string;
    readonly m?: string;
    readonly v?: string;
    readonly wp14?: string;
    readonly wp?: string;
    readonly w10?: string;
    readonly w?: string;
    readonly w14?: string;
    readonly w15?: string;
    readonly wpg?: string;
    readonly wpi?: string;
    readonly wne?: string;
    readonly wps?: string;
    readonly Ignorable?: string;
}> {
    protected readonly xmlKeys = {
        Ignorable: "mc:Ignorable",
        m: "xmlns:m",
        mc: "xmlns:mc",
        o: "xmlns:o",
        r: "xmlns:r",
        v: "xmlns:v",
        w: "xmlns:w",
        w10: "xmlns:w10",
        w14: "xmlns:w14",
        w15: "xmlns:w15",
        wne: "xmlns:wne",
        wp: "xmlns:wp",
        wp14: "xmlns:wp14",
        wpc: "xmlns:wpc",
        wpg: "xmlns:wpg",
        wpi: "xmlns:wpi",
        wps: "xmlns:wps",
    };
}

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
export interface ISettingsOptions {
    /** @deprecated Use compatibility.version instead */
    readonly compatibilityModeVersion?: number;
    /** Enable different headers/footers for even and odd pages */
    readonly evenAndOddHeaders?: boolean;
    /** Enable track changes (revision marking) */
    readonly trackRevisions?: boolean;
    /** Update fields when document is opened */
    readonly updateFields?: boolean;
    /** Compatibility settings for older Word versions */
    readonly compatibility?: ICompatibilityOptions;
    /** Default distance between tab stops in twips */
    readonly defaultTabStop?: number;
    /** Hyphenation settings */
    readonly hyphenation?: IHyphenationOptions;
    /** Controls whether punctuation is compressed at line ends */
    readonly characterSpacingControl?: "compressPunctuation" | "doNotCompress";
    /** Document protection settings */
    readonly documentProtection?: IDocumentProtectionOptions;
    /** Default document view mode */
    readonly view?: "none" | "print" | "outline" | "masterPages" | "normal" | "web";
    /** Default zoom level (percentage) and type */
    readonly zoom?: {
        readonly percent?: number;
        readonly val?: "none" | "fullPage" | "bestFit" | "textFit";
    };
    /** Write protection recommendation (not enforcement) */
    readonly writeProtection?: IWriteProtectionOptions;
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
export interface IDocumentProtectionOptions {
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
export interface IWriteProtectionOptions {
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

/**
 * Options for automatic hyphenation settings.
 *
 * @see {@link Settings}
 */
export interface IHyphenationOptions {
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
    public constructor(options: ISettingsOptions) {
        super("w:settings");
        this.root.push(
            new SettingsAttributes({
                Ignorable: "w14 w15 wp14",
                m: "http://schemas.openxmlformats.org/officeDocument/2006/math",
                mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
                o: "urn:schemas-microsoft-com:office:office",
                r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                v: "urn:schemas-microsoft-com:vml",
                w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                w10: "urn:schemas-microsoft-com:office:word",
                w14: "http://schemas.microsoft.com/office/word/2010/wordml",
                w15: "http://schemas.microsoft.com/office/word/2012/wordml",
                wne: "http://schemas.microsoft.com/office/word/2006/wordml",
                wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                wp14: "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
                wpc: "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
                wpg: "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
                wpi: "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
                wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
            }),
        );

        // https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_trackRevisions_topic_ID0EKXKY.html
        if (options.trackRevisions !== undefined) {
            this.root.push(new OnOffElement("w:trackRevisions", options.trackRevisions));
        }

        // https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_documentProtection_topic_ID0EKCBA.html
        if (options.documentProtection !== undefined) {
            this.root.push(new DocumentProtection(options.documentProtection));
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

        if (options.writeProtection !== undefined) {
            this.root.push(new WriteProtection(options.writeProtection));
        }

        if (options.displayBackgroundShape !== undefined) {
            this.root.push(
                new OnOffElement("w:displayBackgroundShape", options.displayBackgroundShape),
            );
        }

        if (options.embedTrueTypeFonts !== undefined) {
            this.root.push(new OnOffElement("w:embedTrueTypeFonts", options.embedTrueTypeFonts));
        }

        if (options.embedSystemFonts !== undefined) {
            this.root.push(new OnOffElement("w:embedSystemFonts", options.embedSystemFonts));
        }

        if (options.saveSubsetFonts !== undefined) {
            this.root.push(new OnOffElement("w:saveSubsetFonts", options.saveSubsetFonts));
        }

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

        // http://officeopenxml.com/WPSectionFooterReference.php
        // https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_evenAndOddHeaders_topic_ID0ET1WU.html
        if (options.evenAndOddHeaders !== undefined) {
            this.root.push(new OnOffElement("w:evenAndOddHeaders", options.evenAndOddHeaders));
        }

        if (options.updateFields !== undefined) {
            this.root.push(new OnOffElement("w:updateFields", options.updateFields));
        }

        // https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_defaultTabStop_topic_ID0EIXSX.html
        this.root.push(new NumberValueElement("w:defaultTabStop", options.defaultTabStop ?? 420));

        // https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_autoHyphenation_topic_ID0EFUMX.html
        if (options.hyphenation?.autoHyphenation !== undefined) {
            this.root.push(
                new OnOffElement("w:autoHyphenation", options.hyphenation.autoHyphenation),
            );
        }

        // https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_hyphenationZone_topic_ID0ERI3X.html
        if (options.hyphenation?.hyphenationZone !== undefined) {
            this.root.push(
                new NumberValueElement("w:hyphenationZone", options.hyphenation.hyphenationZone),
            );
        }

        // https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_consecutiveHyphenLim_topic_ID0EQ6RX.html
        if (options.hyphenation?.consecutiveHyphenLimit !== undefined) {
            this.root.push(
                new NumberValueElement(
                    "w:consecutiveHyphenLimit",
                    options.hyphenation.consecutiveHyphenLimit,
                ),
            );
        }

        // https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_doNotHyphenateCaps_topic_ID0EW4XX.html
        if (options.hyphenation?.doNotHyphenateCaps !== undefined) {
            this.root.push(
                new OnOffElement("w:doNotHyphenateCaps", options.hyphenation.doNotHyphenateCaps),
            );
        }

        this.root.push(
            new StringValueElement(
                "w:characterSpacingControl",
                options.characterSpacingControl ?? "compressPunctuation",
            ),
        );

        if (options.colorSchemeMapping !== undefined) {
            this.root.push(new ColorSchemeMapping(options.colorSchemeMapping));
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
    }
}

/**
 * Attributes for the documentProtection element.
 *
 * @internal
 */
class DocumentProtectionAttributes extends XmlAttributeComponent<{
    readonly edit?: string;
    readonly enforcement?: string;
    readonly formatting?: string;
    readonly algorithmName?: string;
    readonly hashValue?: string;
    readonly saltValue?: string;
    readonly spinCount?: number;
}> {
    protected readonly xmlKeys = {
        edit: "w:edit",
        enforcement: "w:enforcement",
        formatting: "w:formatting",
        algorithmName: "w:algorithmName",
        hashValue: "w:hashValue",
        saltValue: "w:saltValue",
        spinCount: "w:spinCount",
    };
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
    public constructor(options: IDocumentProtectionOptions) {
        super("w:documentProtection");
        this.root.push(
            new DocumentProtectionAttributes({
                enforcement: "1",
                edit: options.edit,
                formatting:
                    options.formatting !== undefined ? (options.formatting ? "1" : "0") : undefined,
                algorithmName: options.algorithmName,
                hashValue: options.hashValue,
                saltValue: options.saltValue,
                spinCount: options.spinCount,
            }),
        );
    }
}

class ViewAttributes extends XmlAttributeComponent<{ readonly val: string }> {
    protected readonly xmlKeys = { val: "w:val" };
}

class View extends XmlComponent {
    public constructor(val: string) {
        super("w:view");
        this.root.push(new ViewAttributes({ val }));
    }
}

class ZoomAttributes extends XmlAttributeComponent<{
    readonly val?: string;
    readonly percent?: number;
}> {
    protected readonly xmlKeys = { val: "w:val", percent: "w:percent" };
}

class Zoom extends XmlComponent {
    public constructor(options: { val?: string; percent: number }) {
        super("w:zoom");
        this.root.push(new ZoomAttributes(options));
    }
}

class WriteProtectionAttributes extends XmlAttributeComponent<{
    readonly recommended?: string;
    readonly hashValue?: string;
    readonly saltValue?: string;
    readonly spinCount?: number;
    readonly algorithmName?: string;
}> {
    protected readonly xmlKeys = {
        recommended: "w:recommended",
        hashValue: "w:hashValue",
        saltValue: "w:saltValue",
        spinCount: "w:spinCount",
        algorithmName: "w:algorithmName",
    };
}

class WriteProtection extends XmlComponent {
    public constructor(options: IWriteProtectionOptions) {
        super("w:writeProtection");
        this.root.push(
            new WriteProtectionAttributes({
                recommended:
                    options.recommended !== undefined
                        ? options.recommended
                            ? "1"
                            : "0"
                        : undefined,
                hashValue: options.hashValue,
                saltValue: options.saltValue,
                spinCount: options.spinCount,
                algorithmName: options.algorithmName,
            }),
        );
    }
}

class ColorSchemeMappingAttributes extends XmlAttributeComponent<{
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
}> {
    protected readonly xmlKeys = {
        bg1: "w:bg1",
        t1: "w:t1",
        bg2: "w:bg2",
        t2: "w:t2",
        accent1: "w:accent1",
        accent2: "w:accent2",
        accent3: "w:accent3",
        accent4: "w:accent4",
        accent5: "w:accent5",
        accent6: "w:accent6",
        hyperlink: "w:hyperlink",
        followedHyperlink: "w:followedHyperlink",
    };
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
        this.root.push(new ColorSchemeMappingAttributes(options));
    }
}
