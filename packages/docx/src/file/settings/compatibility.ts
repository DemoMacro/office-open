/**
 * Compatibility module for WordprocessingML documents.
 *
 * This module provides compatibility settings that control how Word
 * handles documents created in older versions or other word processors.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-w_compat-1.html
 *
 * @module
 */
import { onOffObj, XmlComponent } from "@file/xml-components";

import { createCompatibilitySetting } from "./compatibility-setting/compatibility-setting";

// <xsd:complexType name="CT_Compat">
// <xsd:sequence>
//   <xsd:element name="useSingleBorderforContiguousCells" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="wpJustification" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="noTabHangInd" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="noLeading" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="spaceForUL" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="noColumnBalance" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="balanceSingleByteDoubleByteWidth" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="noExtraLineSpacing" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="doNotLeaveBackslashAlone" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="ulTrailSpace" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="doNotExpandShiftReturn" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="spacingInWholePoints" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="lineWrapLikeWord6" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="printBodyTextBeforeHeader" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="printColBlack" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="wpSpaceWidth" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="showBreaksInFrames" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="subFontBySize" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="suppressBottomSpacing" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="suppressTopSpacing" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="suppressSpacingAtTopOfPage" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="suppressTopSpacingWP" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="suppressSpBfAfterPgBrk" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="swapBordersFacingPages" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="convMailMergeEsc" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="truncateFontHeightsLikeWP6" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="mwSmallCaps" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="usePrinterMetrics" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="doNotSuppressParagraphBorders" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="wrapTrailSpaces" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="footnoteLayoutLikeWW8" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="shapeLayoutLikeWW8" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="alignTablesRowByRow" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="forgetLastTabAlignment" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="adjustLineHeightInTable" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="autoSpaceLikeWord95" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="noSpaceRaiseLower" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="doNotUseHTMLParagraphAutoSpacing" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="layoutRawTableWidth" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="layoutTableRowsApart" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="useWord97LineBreakRules" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="doNotBreakWrappedTables" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="doNotSnapToGridInCell" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="selectFldWithFirstOrLastChar" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="applyBreakingRules" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="doNotWrapTextWithPunct" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="doNotUseEastAsianBreakRules" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="useWord2002TableStyleRules" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="growAutofit" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="useFELayout" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="useNormalStyleForList" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="doNotUseIndentAsNumberingTabStop" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="useAltKinsokuLineBreakRules" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="allowSpaceOfSameStyleInTable" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="doNotSuppressIndentation" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="doNotAutofitConstrainedTables" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="autofitToFirstFixedWidthCell" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="underlineTabInNumList" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="displayHangulFixedWidth" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="splitPgBreakAndParaMark" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="doNotVertAlignCellWithSp" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="doNotBreakConstrainedForcedTable" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="doNotVertAlignInTxbx" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="useAnsiKerningPairs" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="cachedColBalance" type="CT_OnOff" minOccurs="0"/>
//   <xsd:element name="compatSetting" type="CT_CompatSetting" minOccurs="0" maxOccurs="unbounded"
//   />
// </xsd:sequence>
// </xsd:complexType>

/**
 * Options for configuring document compatibility settings.
 *
 * These settings control how Word processes and displays documents to match
 * behavior from older Word versions or other word processors like WordPerfect.
 *
 * @see {@link Compatibility}
 */
export interface CompatibilityOptions {
  /** Word compatibility mode version (e.g., 15 for Word 2013+) */
  version?: number;
  /** Use Simplified Rules For Table Border Conflicts */
  useSingleBorderforContiguousCells?: boolean;
  /** Emulate WordPerfect 6.x Paragraph Justification */
  wordPerfectJustification?: boolean;
  /** Do Not Create Custom Tab Stop for Hanging Indent */
  noTabStopForHangingIndent?: boolean;
  /** Do Not Add Leading Between Lines of Text */
  noLeading?: boolean;
  /** Add Additional Space Below Baseline For Underlined East Asian Text */
  spaceForUnderline?: boolean;
  /** Do Not Balance Text Columns within a Section */
  noColumnBalance?: boolean;
  /** Balance Single Byte and Double Byte Characters */
  balanceSingleByteDoubleByteWidth?: boolean;
  /** Do Not Center Content on Lines With Exact Line Height */
  noExtraLineSpacing?: boolean;
  /** Convert Backslash To Yen Sign When Entered */
  doNotLeaveBackslashAlone?: boolean;
  /** Underline All Trailing Spaces */
  underlineTrailingSpaces?: boolean;
  /** Don't Justify Lines Ending in Soft Line Break */
  doNotExpandShiftReturn?: boolean;
  /** Only Expand/Condense Text By Whole Points */
  spacingInWholePoints?: boolean;
  /** Emulate Word 6.0 Line Wrapping for East Asian Text */
  lineWrapLikeWord6?: boolean;
  /** Print Body Text before Header/Footer Contents */
  printBodyTextBeforeHeader?: boolean;
  /** Print Colors as Black And White without Dithering */
  printColorsBlack?: boolean;
  /** Space width */
  spaceWidth?: boolean;
  /** Display Page/Column Breaks Present in Frames */
  showBreaksInFrames?: boolean;
  /** Increase Priority Of Font Size During Font Substitution */
  subFontBySize?: boolean;
  /** Ignore Exact Line Height for Last Line on Page */
  suppressBottomSpacing?: boolean;
  /** Ignore Minimum and Exact Line Height for First Line on Page */
  suppressTopSpacing?: boolean;
  /** Ignore Minimum Line Height for First Line on Page */
  suppressSpacingAtTopOfPage?: boolean;
  /** Emulate WordPerfect 5.x Line Spacing */
  suppressTopSpacingWP?: boolean;
  /** Do Not Use Space Before On First Line After a Page Break */
  suppressSpBfAfterPgBrk?: boolean;
  /** Swap Paragraph Borders on Odd Numbered Pages */
  swapBordersFacingPages?: boolean;
  /** Treat Backslash Quotation Delimiter as Two Quotation Marks */
  convertMailMergeEsc?: boolean;
  /** Emulate WordPerfect 6.x Font Height Calculation */
  truncateFontHeightsLikeWP6?: boolean;
  /** Emulate Word 5.x for the Macintosh Small Caps Formatting */
  macWordSmallCaps?: boolean;
  /** Use Printer Metrics To Display Documents */
  usePrinterMetrics?: boolean;
  /** Do Not Suppress Paragraph Borders Next To Frames */
  doNotSuppressParagraphBorders?: boolean;
  /** Line Wrap Trailing Spaces */
  wrapTrailSpaces?: boolean;
  /** Emulate Word 6.x/95/97 Footnote Placement */
  footnoteLayoutLikeWW8?: boolean;
  /** Emulate Word 97 Text Wrapping Around Floating Objects */
  shapeLayoutLikeWW8?: boolean;
  /** Align Table Rows Independently */
  alignTablesRowByRow?: boolean;
  /** Ignore Width of Last Tab Stop When Aligning Paragraph If It Is Not Left Aligned */
  forgetLastTabAlignment?: boolean;
  /** Add Document Grid Line Pitch To Lines in Table Cells */
  adjustLineHeightInTable?: boolean;
  /** Emulate Word 95 Full-Width Character Spacing */
  autoSpaceLikeWord95?: boolean;
  /** Do Not Increase Line Height for Raised/Lowered Text */
  noSpaceRaiseLower?: boolean;
  /** Use Fixed Paragraph Spacing for HTML Auto Setting */
  doNotUseHTMLParagraphAutoSpacing?: boolean;
  /** Ignore Space Before Table When Deciding If Table Should Wrap Floating Object */
  layoutRawTableWidth?: boolean;
  /** Allow Table Rows to Wrap Inline Objects Independently */
  layoutTableRowsApart?: boolean;
  /** Emulate Word 97 East Asian Line Breaking */
  useWord97LineBreakRules?: boolean;
  /** Do Not Allow Floating Tables To Break Across Pages */
  doNotBreakWrappedTables?: boolean;
  /** Do Not Snap to Document Grid in Table Cells with Objects */
  doNotSnapToGridInCell?: boolean;
  /** Select Field When First or Last Character Is Selected */
  selectFieldWithFirstOrLastCharacter?: boolean;
  /** Use Legacy Ethiopic and Amharic Line Breaking Rules */
  applyBreakingRules?: boolean;
  /** Do Not Allow Hanging Punctuation With Character Grid */
  doNotWrapTextWithPunctuation?: boolean;
  /** Do Not Compress Compressible Characters When Using Document Grid */
  doNotUseEastAsianBreakRules?: boolean;
  /** Emulate Word 2002 Table Style Rules */
  useWord2002TableStyleRules?: boolean;
  /** Allow Tables to AutoFit Into Page Margins */
  growAutofit?: boolean;
  /** Do Not Bypass East Asian/Complex Script Layout Code */
  useFELayout?: boolean;
  /** Do Not Automatically Apply List Paragraph Style To Bulleted/Numbered Text */
  useNormalStyleForList?: boolean;
  /** Ignore Hanging Indent When Creating Tab Stop After Numbering */
  doNotUseIndentAsNumberingTabStop?: boolean;
  /** Use Alternate Set of East Asian Line Breaking Rules */
  useAlternateEastAsianLineBreakRules?: boolean;
  /** Allow Contextual Spacing of Paragraphs in Tables */
  allowSpaceOfSameStyleInTable?: boolean;
  /** Do Not Ignore Floating Objects When Calculating Paragraph Indentation */
  doNotSuppressIndentation?: boolean;
  /** Do Not AutoFit Tables To Fit Next To Wrapped Objects */
  doNotAutofitConstrainedTables?: boolean;
  /** Allow Table Columns To Exceed Preferred Widths of Constituent Cells */
  autofitToFirstFixedWidthCell?: boolean;
  /** Underline Following Character Following Numbering */
  underlineTabInNumberingList?: boolean;
  /** Always Use Fixed Width for Hangul Characters */
  displayHangulFixedWidth?: boolean;
  /** Always Move Paragraph Mark to Page after a Page Break */
  splitPgBreakAndParaMark?: boolean;
  /** Don't Vertically Align Cells Containing Floating Objects */
  doNotVerticallyAlignCellWithSp?: boolean;
  /** Don't Break Table Rows Around Floating Tables */
  doNotBreakConstrainedForcedTable?: boolean;
  /** Ignore Vertical Alignment in Textboxes */
  ignoreVerticalAlignmentInTextboxes?: boolean;
  /** Use ANSI Kerning Pairs from Fonts */
  useAnsiKerningPairs?: boolean;
  /** Use Cached Paragraph Information for Column Balancing */
  cachedColumnBalance?: boolean;
  /** Override Table Style Font Size and Justification */
  overrideTableStyleFontSizeAndJustification?: boolean;
  /** Enable OpenType Features */
  enableOpenTypeFeatures?: boolean;
  /** Do Not Flip Mirror Indents */
  doNotFlipMirrorIndents?: boolean;
}

/**
 * Represents compatibility settings in a WordprocessingML document.
 *
 * Compatibility settings control document rendering and layout behavior
 * to match older Word versions or other word processors. This ensures
 * documents maintain consistent appearance across different applications.
 *
 * Reference: http://www.datypic.com/sc/ooxml/e-w_compat-1.html
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Compat">
 *   <xsd:sequence>
 *     <xsd:element name="useSingleBorderforContiguousCells" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="wpJustification" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="noTabHangInd" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="noLeading" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="usePrinterMetrics" type="CT_OnOff" minOccurs="0"/>
 *     <xsd:element name="compatSetting" type="CT_CompatSetting" minOccurs="0" maxOccurs="unbounded"/>
 *     <!-- Additional compatibility elements omitted for brevity -->
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @example
 * ```typescript
 * // Set compatibility mode to Word 2013+
 * new Compatibility({
 *   version: 15,
 * });
 *
 * // Enable specific compatibility options
 * new Compatibility({
 *   version: 15,
 *   usePrinterMetrics: true,
 *   doNotSnapToGridInCell: true,
 * });
 * ```
 */
export class Compatibility extends XmlComponent {
  public constructor(options: CompatibilityOptions) {
    super("w:compat");

    // All individual compat elements first (per XSD CT_Compat sequence)
    if (options.useSingleBorderforContiguousCells) {
      this.root.push(
        onOffObj("w:useSingleBorderforContiguousCells", options.useSingleBorderforContiguousCells),
      );
    }

    if (options.wordPerfectJustification) {
      this.root.push(onOffObj("w:wpJustification", options.wordPerfectJustification));
    }

    if (options.noTabStopForHangingIndent) {
      this.root.push(onOffObj("w:noTabHangInd", options.noTabStopForHangingIndent));
    }

    if (options.noLeading) {
      this.root.push(onOffObj("w:noLeading", options.noLeading));
    }

    if (options.spaceForUnderline) {
      this.root.push(onOffObj("w:spaceForUL", options.spaceForUnderline));
    }

    if (options.noColumnBalance) {
      this.root.push(onOffObj("w:noColumnBalance", options.noColumnBalance));
    }

    if (options.balanceSingleByteDoubleByteWidth) {
      this.root.push(
        onOffObj("w:balanceSingleByteDoubleByteWidth", options.balanceSingleByteDoubleByteWidth),
      );
    }

    if (options.noExtraLineSpacing) {
      this.root.push(onOffObj("w:noExtraLineSpacing", options.noExtraLineSpacing));
    }

    if (options.doNotLeaveBackslashAlone) {
      this.root.push(onOffObj("w:doNotLeaveBackslashAlone", options.doNotLeaveBackslashAlone));
    }

    if (options.underlineTrailingSpaces) {
      this.root.push(onOffObj("w:ulTrailSpace", options.underlineTrailingSpaces));
    }

    if (options.doNotExpandShiftReturn) {
      this.root.push(onOffObj("w:doNotExpandShiftReturn", options.doNotExpandShiftReturn));
    }

    if (options.spacingInWholePoints) {
      this.root.push(onOffObj("w:spacingInWholePoints", options.spacingInWholePoints));
    }

    if (options.lineWrapLikeWord6) {
      this.root.push(onOffObj("w:lineWrapLikeWord6", options.lineWrapLikeWord6));
    }

    if (options.printBodyTextBeforeHeader) {
      this.root.push(onOffObj("w:printBodyTextBeforeHeader", options.printBodyTextBeforeHeader));
    }

    if (options.printColorsBlack) {
      this.root.push(onOffObj("w:printColBlack", options.printColorsBlack));
    }

    if (options.spaceWidth) {
      this.root.push(onOffObj("w:wpSpaceWidth", options.spaceWidth));
    }

    if (options.showBreaksInFrames) {
      this.root.push(onOffObj("w:showBreaksInFrames", options.showBreaksInFrames));
    }

    if (options.subFontBySize) {
      this.root.push(onOffObj("w:subFontBySize", options.subFontBySize));
    }

    if (options.suppressBottomSpacing) {
      this.root.push(onOffObj("w:suppressBottomSpacing", options.suppressBottomSpacing));
    }

    if (options.suppressTopSpacing) {
      this.root.push(onOffObj("w:suppressTopSpacing", options.suppressTopSpacing));
    }

    if (options.suppressSpacingAtTopOfPage) {
      this.root.push(onOffObj("w:suppressSpacingAtTopOfPage", options.suppressSpacingAtTopOfPage));
    }

    if (options.suppressTopSpacingWP) {
      this.root.push(onOffObj("w:suppressTopSpacingWP", options.suppressTopSpacingWP));
    }

    if (options.suppressSpBfAfterPgBrk) {
      this.root.push(onOffObj("w:suppressSpBfAfterPgBrk", options.suppressSpBfAfterPgBrk));
    }

    if (options.swapBordersFacingPages) {
      this.root.push(onOffObj("w:swapBordersFacingPages", options.swapBordersFacingPages));
    }

    if (options.convertMailMergeEsc) {
      this.root.push(onOffObj("w:convMailMergeEsc", options.convertMailMergeEsc));
    }

    if (options.truncateFontHeightsLikeWP6) {
      this.root.push(onOffObj("w:truncateFontHeightsLikeWP6", options.truncateFontHeightsLikeWP6));
    }

    if (options.macWordSmallCaps) {
      this.root.push(onOffObj("w:mwSmallCaps", options.macWordSmallCaps));
    }

    if (options.usePrinterMetrics) {
      this.root.push(onOffObj("w:usePrinterMetrics", options.usePrinterMetrics));
    }

    if (options.doNotSuppressParagraphBorders) {
      this.root.push(
        onOffObj("w:doNotSuppressParagraphBorders", options.doNotSuppressParagraphBorders),
      );
    }

    if (options.wrapTrailSpaces) {
      this.root.push(onOffObj("w:wrapTrailSpaces", options.wrapTrailSpaces));
    }

    if (options.footnoteLayoutLikeWW8) {
      this.root.push(onOffObj("w:footnoteLayoutLikeWW8", options.footnoteLayoutLikeWW8));
    }

    if (options.shapeLayoutLikeWW8) {
      this.root.push(onOffObj("w:shapeLayoutLikeWW8", options.shapeLayoutLikeWW8));
    }

    if (options.alignTablesRowByRow) {
      this.root.push(onOffObj("w:alignTablesRowByRow", options.alignTablesRowByRow));
    }

    if (options.forgetLastTabAlignment) {
      this.root.push(onOffObj("w:forgetLastTabAlignment", options.forgetLastTabAlignment));
    }

    if (options.adjustLineHeightInTable) {
      this.root.push(onOffObj("w:adjustLineHeightInTable", options.adjustLineHeightInTable));
    }

    if (options.autoSpaceLikeWord95) {
      this.root.push(onOffObj("w:autoSpaceLikeWord95", options.autoSpaceLikeWord95));
    }

    if (options.noSpaceRaiseLower) {
      this.root.push(onOffObj("w:noSpaceRaiseLower", options.noSpaceRaiseLower));
    }

    if (options.doNotUseHTMLParagraphAutoSpacing) {
      this.root.push(
        onOffObj("w:doNotUseHTMLParagraphAutoSpacing", options.doNotUseHTMLParagraphAutoSpacing),
      );
    }

    if (options.layoutRawTableWidth) {
      this.root.push(onOffObj("w:layoutRawTableWidth", options.layoutRawTableWidth));
    }

    if (options.layoutTableRowsApart) {
      this.root.push(onOffObj("w:layoutTableRowsApart", options.layoutTableRowsApart));
    }

    if (options.useWord97LineBreakRules) {
      this.root.push(onOffObj("w:useWord97LineBreakRules", options.useWord97LineBreakRules));
    }

    if (options.doNotBreakWrappedTables) {
      this.root.push(onOffObj("w:doNotBreakWrappedTables", options.doNotBreakWrappedTables));
    }

    if (options.doNotSnapToGridInCell) {
      this.root.push(onOffObj("w:doNotSnapToGridInCell", options.doNotSnapToGridInCell));
    }

    if (options.selectFieldWithFirstOrLastCharacter) {
      this.root.push(
        onOffObj("w:selectFldWithFirstOrLastChar", options.selectFieldWithFirstOrLastCharacter),
      );
    }

    if (options.applyBreakingRules) {
      this.root.push(onOffObj("w:applyBreakingRules", options.applyBreakingRules));
    }

    if (options.doNotWrapTextWithPunctuation) {
      this.root.push(onOffObj("w:doNotWrapTextWithPunct", options.doNotWrapTextWithPunctuation));
    }

    if (options.doNotUseEastAsianBreakRules) {
      this.root.push(
        onOffObj("w:doNotUseEastAsianBreakRules", options.doNotUseEastAsianBreakRules),
      );
    }

    if (options.useWord2002TableStyleRules) {
      this.root.push(onOffObj("w:useWord2002TableStyleRules", options.useWord2002TableStyleRules));
    }

    if (options.growAutofit) {
      this.root.push(onOffObj("w:growAutofit", options.growAutofit));
    }

    if (options.useFELayout) {
      this.root.push(onOffObj("w:useFELayout", options.useFELayout));
    }

    if (options.useNormalStyleForList) {
      this.root.push(onOffObj("w:useNormalStyleForList", options.useNormalStyleForList));
    }

    if (options.doNotUseIndentAsNumberingTabStop) {
      this.root.push(
        onOffObj("w:doNotUseIndentAsNumberingTabStop", options.doNotUseIndentAsNumberingTabStop),
      );
    }

    if (options.useAlternateEastAsianLineBreakRules) {
      this.root.push(
        onOffObj("w:useAltKinsokuLineBreakRules", options.useAlternateEastAsianLineBreakRules),
      );
    }

    if (options.allowSpaceOfSameStyleInTable) {
      this.root.push(
        onOffObj("w:allowSpaceOfSameStyleInTable", options.allowSpaceOfSameStyleInTable),
      );
    }

    if (options.doNotSuppressIndentation) {
      this.root.push(onOffObj("w:doNotSuppressIndentation", options.doNotSuppressIndentation));
    }

    if (options.doNotAutofitConstrainedTables) {
      this.root.push(
        onOffObj("w:doNotAutofitConstrainedTables", options.doNotAutofitConstrainedTables),
      );
    }

    if (options.autofitToFirstFixedWidthCell) {
      this.root.push(
        onOffObj("w:autofitToFirstFixedWidthCell", options.autofitToFirstFixedWidthCell),
      );
    }

    if (options.underlineTabInNumberingList) {
      this.root.push(onOffObj("w:underlineTabInNumList", options.underlineTabInNumberingList));
    }

    if (options.displayHangulFixedWidth) {
      this.root.push(onOffObj("w:displayHangulFixedWidth", options.displayHangulFixedWidth));
    }

    if (options.splitPgBreakAndParaMark) {
      this.root.push(onOffObj("w:splitPgBreakAndParaMark", options.splitPgBreakAndParaMark));
    }

    if (options.doNotVerticallyAlignCellWithSp) {
      this.root.push(
        onOffObj("w:doNotVertAlignCellWithSp", options.doNotVerticallyAlignCellWithSp),
      );
    }

    if (options.doNotBreakConstrainedForcedTable) {
      this.root.push(
        onOffObj("w:doNotBreakConstrainedForcedTable", options.doNotBreakConstrainedForcedTable),
      );
    }

    if (options.ignoreVerticalAlignmentInTextboxes) {
      this.root.push(
        onOffObj("w:doNotVertAlignInTxbx", options.ignoreVerticalAlignmentInTextboxes),
      );
    }

    if (options.useAnsiKerningPairs) {
      this.root.push(onOffObj("w:useAnsiKerningPairs", options.useAnsiKerningPairs));
    }

    if (options.cachedColumnBalance) {
      this.root.push(onOffObj("w:cachedColBalance", options.cachedColumnBalance));
    }

    // compatSetting must come last per XSD CT_Compat sequence
    if (options.version) {
      this.root.push(createCompatibilitySetting("compatibilityMode", options.version));
    }

    if (options.overrideTableStyleFontSizeAndJustification) {
      this.root.push(createCompatibilitySetting("overrideTableStyleFontSizeAndJustification", 1));
    }

    if (options.enableOpenTypeFeatures) {
      this.root.push(createCompatibilitySetting("enableOpenTypeFeatures", 1));
    }

    if (options.doNotFlipMirrorIndents) {
      this.root.push(createCompatibilitySetting("doNotFlipMirrorIndents", 1));
    }
  }
}
