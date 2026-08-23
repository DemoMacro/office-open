/**
 * Table of Contents Properties module.
 *
 * This module defines configuration options for table of contents generation,
 * including field switches and style mappings.
 *
 * Reference: http://officeopenxml.com/WPtableOfContents.php
 *
 * @module
 */

import type { RunPropertiesOptions } from "@parts/paragraph/run/properties";
import type { SectionChild } from "@shared/section";

/**
 * Represents a style-to-level mapping for table of contents entries.
 *
 * StyleLevel associates a paragraph style name with a TOC level, allowing
 * custom styles to be included in the table of contents at specific levels.
 *
 * @publicApi
 */
export class StyleLevel {
  /** The name of the paragraph style. */
  public styleName: string;
  /** The TOC level (1-9) to assign to this style. */
  public level: number;

  public constructor(styleName: string, level: number) {
    this.styleName = styleName;
    this.level = level;
  }
}

/**
 * Options for configuring a Table of Contents.
 *
 * These options control which content is included in the TOC and how it is formatted.
 * Options correspond to field switches in the TOC field code.
 *
 * Reference:
 * - https://www.ecma-international.org/publications/standards/Ecma-376.htm (Part 1, Page 1251)
 * - http://officeopenxml.com/WPtableOfContents.php
 */
export interface TableOfContentsOptions {
  /**
   * \a option — include captioned items (SEQ-numbered), omitting the caption
   * label and number. Argument is the caption-label identifier. Use \c
   * (captionLabelIncludingNumbers) to keep labels and numbers.
   */
  captionLabel?: string;

  /**
   * \b option - Includes entries only from the portion of the document marked by
   * the bookmark named by text in this switch's field-argument.
   */
  entriesFromBookmark?: string;

  /**
   * \c option — include figures/tables/charts numbered by a SEQ field, with
   * caption labels and numbers. Argument is the caption-label/sequence
   * identifier and must match the SEQ field's identifier.
   */
  captionLabelIncludingNumbers?: string;

  /**
   * \d option - When used with \s, the text in this switch's field-argument defines
   * the separator between sequence and page numbers. The default separator is a hyphen (-).
   */
  sequenceAndPageNumbersSeparator?: string;

  /**
   * \f option - Includes only those TC fields whose identifier exactly matches the
   * text in this switch's field-argument (which is typically a letter).
   */
  tcFieldIdentifier?: string;

  /**
   * \h option - Makes the table of contents entries hyperlinks.
   */
  hyperlink?: boolean;

  /**
   * \l option — include TC fields that assign entries to levels in the
   * argument range `startLevel-endLevel` (integers, start ≤ end). TC fields
   * assigning lower levels are skipped.
   */
  tcFieldLevelRange?: string;

  /**
   * \n option - Without field-argument, omits page numbers from the table of contents.
   * Page numbers are omitted from all levels unless a range of entry levels is specified by
   * text in this switch's field-argument. A range is specified as for \l.
   */
  pageNumbersEntryLevelsRange?: string;

  /**
   * \o option — build entries from paragraphs in built-in heading styles.
   * Argument is a level range as for \l (e.g. "1-3"; 1 = Heading1); empty or
   * omitted lists all heading levels used in the document.
   */
  headingStyleRange?: string;

  /**
   * \p option - Text in this switch's field-argument specifies a sequence of characters
   * that separate an entry and its page number. The default is a tab with leader dots.
   */
  entryAndPageNumberSeparator?: string;

  /**
   * \s option - For entries numbered with a SEQ field, adds a prefix to the page number.
   * The prefix depends on the type of entry; the field-argument text must match the
   * identifier in the SEQ field.
   */
  seqFieldIdentifierForPrefix?: string;

  /**
   * \t option — include paragraphs in non-heading styles, each mapped to a TOC
   * level. Combinable with \o (headingStyleRange).
   */
  stylesWithLevels?: StyleLevel[];

  /**
   * \u Uses the applied paragraph outline level.
   */
  useAppliedParagraphOutlineLevel?: boolean;

  /**
   * \w Preserves tab entries within table entries.
   */
  preserveTabInEntries?: boolean;

  /**
   * \x Preserves newline characters within table entries.
   */
  preserveNewLineInEntries?: boolean;

  /**
   * \z Hides tab leader and page numbers in web page view (§17.18.102).
   */
  hideTabAndPageNumbersInWebView?: boolean;

  /**
   * Rendered TOC entries — the paragraphs between the field's `separate` and
   * `end` markers, round-tripped structurally (HYPERLINK field, tab leader,
   * PAGEREF page number) so Office/WPS display them without regenerating.
   * Omit on fresh generation (field is emitted `dirty`).
   */
  entries?: SectionChild[];

  /**
   * Run properties carried by the field's begin/instruction/separate control
   * runs, as serialized `<w:rPr>…</w:rPr>` XML (round-trip only; Word parks
   * explicit style overrides — e.g. `b w:val="0"` — on these invisible runs).
   */
  rPrXml?: string;

  /**
   * Run properties of the field's closing `end` run (round-trip only).
   * `""` means the end run carried no rPr (stays bare rather than inheriting
   * the control rPr); `undefined` means it matched the control runs' rPr.
   */
  endRPrXml?: string;

  /**
   * Emit the TOC as a bare complex field instead of the sdt content control
   * Word normally wraps it in. Set by parsing when the source carried no
   * `w:sdt` wrapper, so the document round-trips in its original form.
   */
  bare?: boolean;

  /**
   * The field-end marker lives in a following body paragraph that round-trips
   * on its own — stringify must not inject another end run into the last
   * entry (round-trip only).
   */
  endInBody?: boolean;

  /**
   * Verbatim `<w:r>…</w:r>` XML of the field's begin→separate control runs
   * (round-trip only; Word splits the instruction across runs and the split
   * must survive).
   */
  headRunsXml?: string;

  /**
   * Run properties of the SDT start mark — sdtPr's leading w:rPr element
   * (round-trip only; Word parks the TOC style's font overrides there).
   */
  runProperties?: RunPropertiesOptions;

  /** SDT identifier (w:id) of the TOC content control (round-trip only). */
  id?: number;

  /** Emit w:docPartUnique inside sdtPr's w:docPartObj (round-trip only). */
  docPartUnique?: boolean;

  /**
   * Run properties of the SDT end mark (w:sdtEndPr). An empty object means
   * the source carried a bare `<w:sdtEndPr/>` (round-trip only).
   */
  endProperties?: RunPropertiesOptions;

  /**
   * Block content the sdtContent carries before the TOC field — the TOC
   * heading paragraph, a wrapping bookmarkStart (round-trip only).
   */
  leading?: SectionChild[];

  /**
   * Block content the sdtContent carries after the TOC field — the closing
   * bookmarkEnd of a wrapping bookmark (round-trip only).
   */
  trailing?: SectionChild[];
}
