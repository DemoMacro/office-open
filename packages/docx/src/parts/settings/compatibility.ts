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
